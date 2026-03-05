"""
산업안전보건관리비 진위여부 확인 시스템 (웹 UI)
================================================
Flask 기반 웹 애플리케이션 — 브라우저에서 접속하여 사용

실행: python app_web.py
접속: http://localhost:5000

필수: pip install flask pandas openpyxl PyMuPDF pyzbar opencv-python-headless
              pillow zxing-cpp requests
"""

import os, sys, re, json, time, random, threading, queue, logging, io, uuid
from pathlib import Path
from datetime import datetime

from flask import (
    Flask, render_template_string, request, jsonify,
    Response, send_file, stream_with_context,
)
import pandas as pd

# ──────────────────────────────────────────────
# 앱 전역 상태
# ──────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = "safety-ocr-2026"

# SSE 이벤트 큐 (브라우저로 실시간 전송)
_clients: dict[str, queue.Queue] = {}

# 작업 상태
_state = {
    "df": None,           # pandas DataFrame (결과)
    "running": False,     # 작업 진행 중 여부
    "stop_requested": False,  # 중지 요청 플래그
    "folder": "",         # 현재 폴더
    "progress": 0,
    "total": 0,
    "message": "대기 중",
}


# ──────────────────────────────────────────────
# 로깅 → SSE 브로드캐스트
# ──────────────────────────────────────────────
class SSELogHandler(logging.Handler):
    def emit(self, record):
        msg = self.format(record)
        _broadcast("log", msg)


_sse_handler = SSELogHandler()
_sse_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
logging.root.addHandler(_sse_handler)
logging.root.setLevel(logging.INFO)
log = logging.getLogger("app_web")


def _broadcast(event: str, data):
    """모든 SSE 클라이언트에게 이벤트 전송"""
    payload = json.dumps(data, ensure_ascii=False) if not isinstance(data, str) else data
    dead = []
    for cid, q in _clients.items():
        try:
            q.put_nowait((event, payload))
        except queue.Full:
            dead.append(cid)
    for cid in dead:
        _clients.pop(cid, None)


def _send_progress(current: int, total: int, message: str = ""):
    _state["progress"] = current
    _state["total"] = total
    _state["message"] = message or f"{current}/{total}"
    _broadcast("progress", {"current": current, "total": total, "message": _state["message"]})


def _send_table():
    """DataFrame → JSON 으로 테이블 갱신 이벤트"""
    if _state["df"] is None:
        return
    rows = []
    for i, (_, r) in enumerate(_state["df"].iterrows()):
        amt = r.get("합계금액", "")
        try:
            amt = f"{int(str(amt).replace(',', '')):,}"
        except (ValueError, TypeError):
            amt = str(amt)
        rows.append({
            "no": i + 1,
            "filename": str(r.get("파일명", "")),
            "page": str(r.get("인식페이지", "")),
            "approval": str(r.get("승인번호", "")),
            "supplier": str(r.get("공급자번호", "")),
            "date": str(r.get("작성일자", "")),
            "amount": amt,
            "verdict": str(r.get("진위여부", "")),
            "usage": str(r.get("사용용도", "")),
            "remark": str(r.get("비고", "")),
        })
    _broadcast("table", rows)


# ──────────────────────────────────────────────
# 지연 임포트
# ──────────────────────────────────────────────
def _import_qr():
    import df_qr_batch as m
    return m


def _import_verify():
    import hometax_verify as m
    return m


# ──────────────────────────────────────────────
# 증빙자료 분류 데이터 (data.xlsx 로드)
# ──────────────────────────────────────────────
def _load_evidence_data():
    """data.xlsx → 사용용도별 증빙자료 구조화"""
    data_path = Path(__file__).parent / "data.xlsx"
    if not data_path.exists():
        return {}

    df = pd.read_excel(str(data_path), engine="openpyxl")
    result = {}

    for _, row in df.iterrows():
        category = str(row.get("사용용도", "")).strip()
        sub = str(row.get("Unnamed: 1", "")).strip()
        if category == "nan" or not category:
            continue

        key = f"{category} ({sub})" if sub and sub != "nan" else category

        if key not in result:
            result[key] = {"증빙자료1": [], "증빙자료2": [], "증빙자료3": []}

        for col_name in ["증빙자료1", "증빙자료2", "증빙자료3"]:
            val = str(row.get(col_name, "")).strip()
            if val and val != "nan" and val not in result[key][col_name]:
                result[key][col_name].append(val)

        # Unnamed: 5 (4번째 증빙) → 증빙자료3에 합침
        extra = str(row.get("Unnamed: 5", "")).strip()
        if extra and extra != "nan" and extra not in result[key].get("증빙자료3", []):
            result[key]["증빙자료3"].append(extra)

    return result


_evidence_data = _load_evidence_data()


# ──────────────────────────────────────────────
# 사용용도 자동 분류 & 누락 증빙 확인
# ──────────────────────────────────────────────
_USAGE_KEYWORDS = {
    "보호구": ["안전용품", "보호구", "안전벨트", "경보기", "방염복", "안전화", "안전모",
               "안전장갑", "보호장비", "활선", "혈압계", "안전장구", "철물공구"],
    "안전시설비": ["버킷", "플레이트", "안전시설", "시설물", "안전표지", "안전난간",
                   "보강", "설치공사"],
    "근로자 건강장해예방비": ["생수", "음료", "건강", "식수", "샘물", "정수", "식음료",
                             "냉온수", "얼음"],
    "안전보건교육비": ["교육", "위험성평가담당자교육", "안전교육", "보건교육", "수료",
                       "위험성평가 교육", "특별안전교육"],
    "안전보건진단비": ["안전진단", "보건진단", "정밀진단"],
    "기술지도비": ["기술지도", "컨설팅", "재해예방기술"],
    "안전관리자 임금 등": ["임금", "급여", "인건비", "관리자"],
    "위험성평가 등에 따른 소요비용": ["위험성평가표", "위험성평가 결과"],
}

# 사용용도별 증빙자료1 외 추가 필요 증빙
_ADDITIONAL_EVIDENCE = {
    "보호구": ["거래명세표", "지급대장(안전장구 착용 사진 포함)"],
    "안전시설비": ["거래명세표"],
    "근로자 건강장해예방비": ["입증자료(구입사진 등)"],
    "안전보건교육비": ["수료증 또는 거래명세표"],
    "안전보건진단비": ["계약서", "유자격확인자료"],
    "기술지도비": ["계약서", "결과보고서"],
    "위험성평가 등에 따른 소요비용": ["노사협의체 협의록", "위험성평가표",
                                      "위험성평가 결과 검토서(발주부서 승인분)"],
}

# 다중 페이지 PDF 내 포함 증빙 감지 키워드
_DOC_TYPE_KEYWORDS = {
    "거래명세표": ["거래명세", "거래명세서", "거래명세표"],
    "지급대장": ["지급대장", "보호구지급", "개인보호구지급"],
    "사진": ["사진대지", "사진설명", "사진 설명"],
    "수료증": ["수료증", "수료", "이수증"],
    "계약서": ["계약서", "용역계약", "도급계약"],
    "결과보고서": ["결과보고서", "결과 보고서"],
    "입증자료": ["구입사진", "입증자료", "납품사진"],
    "노사협의체 협의록": ["노사협의", "협의록"],
    "위험성평가표": ["위험성평가표", "위험성 평가표"],
}


def _classify_and_check(df, folder):
    """3단계: 사용용도 분류 + 누락 증빙 확인"""
    import fitz

    if "사용용도" not in df.columns:
        df["사용용도"] = ""
    if "비고" not in df.columns:
        df["비고"] = ""

    for idx, row in df.iterrows():
        filename = str(row.get("파일명", ""))
        pdf_path = Path(folder) / filename

        # PDF 전체 텍스트 추출
        full_text = filename
        page_texts = []
        if pdf_path.exists():
            try:
                doc = fitz.open(str(pdf_path))
                for page in doc:
                    pt = page.get_text()
                    page_texts.append(pt)
                    full_text += " " + pt
                doc.close()
            except Exception:
                pass

        # ── 사용용도 분류 ──
        matched = ""
        for category, keywords in _USAGE_KEYWORDS.items():
            for kw in keywords:
                if kw in full_text:
                    matched = category
                    break
            if matched:
                break
        df.at[idx, "사용용도"] = matched or "분류필요"

        # ── 누락 증빙 확인 ──
        if not matched or matched not in _ADDITIONAL_EVIDENCE:
            df.at[idx, "비고"] = "담당자 세부별도 확인 필요"
            continue

        needed = _ADDITIONAL_EVIDENCE[matched]

        # 다중 페이지 PDF에서 포함된 증빙 감지
        found_in_pdf = set()
        for pt in page_texts:
            for doc_type, doc_kws in _DOC_TYPE_KEYWORDS.items():
                for dk in doc_kws:
                    if dk in pt:
                        found_in_pdf.add(doc_type)

        # 누락된 증빙 판단
        missing = []
        for req in needed:
            # req에 "또는"이 포함된 경우 어느 하나만 있으면 OK
            req_parts = [r.strip() for r in req.split("또는")]
            found = False
            for rp in req_parts:
                for doc_type in found_in_pdf:
                    if doc_type in rp or rp in doc_type:
                        found = True
                        break
                # 사진 관련 체크
                if "사진" in rp and "사진" in found_in_pdf:
                    found = True
                if found:
                    break
            if not found:
                missing.append(req)

        if missing:
            df.at[idx, "비고"] = f"담당자 세부별도 확인 필요 ({', '.join(missing)})"
        else:
            df.at[idx, "비고"] = "증빙 완비"

    return df


# ══════════════════════════════════════════════
# API 엔드포인트
# ══════════════════════════════════════════════

# ── SSE 스트림 ────────────────────────────────
@app.route("/stream")
def stream():
    cid = str(uuid.uuid4())
    q: queue.Queue = queue.Queue(maxsize=200)
    _clients[cid] = q

    def gen():
        try:
            while True:
                try:
                    event, data = q.get(timeout=30)
                    yield f"event: {event}\ndata: {data}\n\n"
                except queue.Empty:
                    yield ": keepalive\n\n"   # 연결 유지
        except GeneratorExit:
            _clients.pop(cid, None)

    return Response(stream_with_context(gen()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ── 상태 조회 ──────────────────────────────────
@app.route("/api/status")
def api_status():
    df = _state["df"]
    stats = {}
    if df is not None and not df.empty:
        stats = {
            "total":   len(df),
            "qr_ok":   int((df["승인번호"].notna() & ~df["승인번호"].isin(["인식 불가"])).sum()) if "승인번호" in df.columns else 0,
            "ok":      int((df.get("진위여부") == "O").sum()) if "진위여부" in df.columns else 0,
            "fail":    int((df.get("진위여부") == "X").sum()) if "진위여부" in df.columns else 0,
            "unknown": int((df.get("진위여부") == "확인불가").sum()) if "진위여부" in df.columns else 0,
        }
    return jsonify({
        "running":  _state["running"],
        "folder":   _state["folder"],
        "progress": _state["progress"],
        "totalProg":_state["total"],
        "message":  _state["message"],
        "stats":    stats,
    })


# ── 1단계: QR 추출 ────────────────────────────
@app.route("/api/extract", methods=["POST"])
def api_extract():
    if _state["running"]:
        return jsonify({"error": "이미 작업 진행 중"}), 409

    folder = request.json.get("folder", "").strip()
    if not folder or not Path(folder).is_dir():
        return jsonify({"error": "유효한 폴더 경로를 입력하세요"}), 400

    pdf_files = sorted(Path(folder).glob("*.pdf"))
    if not pdf_files:
        return jsonify({"error": "폴더에 PDF 파일이 없습니다"}), 400

    _state["folder"] = folder
    threading.Thread(target=_worker_extract, args=(folder, pdf_files), daemon=True).start()
    return jsonify({"ok": True, "count": len(pdf_files)})


def _worker_extract(folder: str, pdf_files: list):
    _state["running"] = True
    _broadcast("status", "running")
    try:
        qr = _import_qr()
        records = []
        total = len(pdf_files)

        for i, pdf_path in enumerate(pdf_files):
            log.info(f"[{i+1}/{total}] {pdf_path.name}")
            _send_progress(i, total, f"QR 스캔 중... {i+1}/{total}")
            record = qr.process_pdf(str(pdf_path))
            records.append(record)

        col_order = ["파일명", "인식페이지", "승인번호", "공급자번호", "작성일자", "합계금액", "원본QR"]
        _state["df"] = pd.DataFrame(records, columns=col_order)

        out = str(Path(folder) / "results.xlsx")
        _state["df"].to_excel(out, index=False, engine="openpyxl")

        success = (_state["df"]["승인번호"].notna() & ~_state["df"]["승인번호"].isin(["인식 불가"])).sum()
        log.info(f"QR 추출 완료: 성공 {success} / 전체 {total}")
        _send_progress(total, total, "QR 추출 완료!")
        _send_table()

    except Exception as e:
        log.error(f"QR 추출 오류: {e}")
        _broadcast("error", str(e))
    finally:
        _state["running"] = False
        _broadcast("status", "idle")


# ── 전체 자동 (1단계→2단계) ───────────────────
@app.route("/api/full", methods=["POST"])
def api_full():
    if _state["running"]:
        return jsonify({"error": "이미 작업 진행 중"}), 409

    folder = request.json.get("folder", "").strip()
    if not folder or not Path(folder).is_dir():
        return jsonify({"error": "유효한 폴더 경로를 입력하세요"}), 400

    pdf_files = sorted(Path(folder).glob("*.pdf"))
    if not pdf_files:
        return jsonify({"error": "폴더에 PDF 파일이 없습니다"}), 400

    _state["folder"] = folder
    threading.Thread(target=_worker_full, args=(folder, pdf_files), daemon=True).start()
    return jsonify({"ok": True, "count": len(pdf_files)})


def _worker_full(folder: str, pdf_files: list):
    """1단계 QR 추출 → 2단계 진위 확인 연속 실행"""
    # ── 1단계: QR 추출 ──
    _state["running"] = True
    _state["stop_requested"] = False
    _broadcast("status", "running")
    try:
        qr = _import_qr()
        records = []
        total = len(pdf_files)

        for i, pdf_path in enumerate(pdf_files):
            if _state["stop_requested"]:
                log.info(f"1단계 중지됨 ({i}/{total})")
                break
            log.info(f"[1단계 {i+1}/{total}] {pdf_path.name}")
            _send_progress(i, total, f"1단계: QR 스캔 중... {i+1}/{total}")
            record = qr.process_pdf(str(pdf_path))
            records.append(record)

        col_order = ["파일명", "인식페이지", "승인번호", "공급자번호", "작성일자", "합계금액", "원본QR"]
        _state["df"] = pd.DataFrame(records, columns=col_order)

        out = str(Path(folder) / "results.xlsx")
        _state["df"].to_excel(out, index=False, engine="openpyxl")

        success = (_state["df"]["승인번호"].notna() & ~_state["df"]["승인번호"].isin(["인식 불가"])).sum()
        log.info(f"1단계 완료: QR 인식 {success} / 전체 {len(records)}")
        _send_progress(total, total, "1단계 완료! 2단계 진위 확인 시작...")
        _send_table()

        if _state["stop_requested"]:
            log.info("중지 요청으로 2단계 건너뜀. 이어서 처리로 재개 가능합니다.")
            _state["running"] = False
            _broadcast("status", "idle")
            return

    except Exception as e:
        log.error(f"QR 추출 오류: {e}")
        _broadcast("error", str(e))
        _state["running"] = False
        _broadcast("status", "idle")
        return

    # ── 2단계: 진위 확인 ──
    _run_verify_stage(folder)


def _run_verify_stage(folder: str):
    """2단계 진위 확인 (full / resume 공용)"""
    import urllib3
    urllib3.disable_warnings()
    try:
        verify = _import_verify()
        df = _state["df"]

        if "진위여부" not in df.columns:
            df["진위여부"] = ""

        total = len(df)
        done_count = 0
        for idx, row in df.iterrows():
            i = idx + 1
            if _state["stop_requested"]:
                log.info(f"2단계 중지됨 ({i-1}/{total}). 중간결과 저장 중...")
                break

            existing = str(row.get("진위여부", "")).strip()
            if existing in ("O", "X"):
                done_count += 1
                _send_progress(i, total, f"2단계: 진위 확인 중... {i}/{total} (기확인 스킵)")
                continue

            if str(row.get("승인번호", "")).startswith("인식"):
                df.at[idx, "진위여부"] = "확인불가"
                _send_progress(i, total, f"2단계: 진위 확인 중... {i}/{total}")
                _send_table()
                continue

            log.info(f"[2단계 {i}/{total}] {row['파일명']}")
            _send_progress(i - 1, total, f"2단계: 진위 확인 중... {i}/{total}")

            row_data = {
                "승인번호_원본": str(row.get("승인번호", "")),
                "승인번호_정제": verify.clean_approval_number(str(row.get("승인번호", ""))),
                "공급자번호":    verify.clean_number(row.get("공급자번호", "")),
                "작성일자_raw":  verify.clean_date(str(row.get("작성일자", ""))),
                "합계금액":      verify.clean_number(row.get("합계금액", "")),
                "수급자번호":    "",
            }

            result = verify.verify_via_api(row_data) or "확인불가"
            df.at[idx, "진위여부"] = result
            log.info(f"  판정: {result}")
            _send_table()

            # 10건마다 자동 중간 저장
            if (i - done_count) % 10 == 0 and folder:
                df.to_excel(str(Path(folder) / "verification_final.xlsx"), index=False, engine="openpyxl")
                log.info(f"  [자동저장] {i}/{total}")

            time.sleep(random.uniform(2, 5))

        # 최종 저장
        if folder:
            out = str(Path(folder) / "verification_final.xlsx")
            df.to_excel(out, index=False, engine="openpyxl")
            log.info(f"저장 → {out}")

        ok  = (df["진위여부"] == "O").sum()
        ng  = (df["진위여부"] == "X").sum()
        unk = (df["진위여부"] == "확인불가").sum()

        if _state["stop_requested"]:
            log.info(f"중지됨. 현재까지 O:{ok}  X:{ng}  ?:{unk} — 이어서 처리로 재개 가능")
            _send_progress(0, 0, "중지됨 — 이어서 처리 가능")
        else:
            # ── 3단계: 증빙 분류 & 누락 확인 ──
            log.info("3단계: 사용용도 분류 및 증빙 누락 확인 중...")
            _send_progress(total, total, "3단계: 증빙 확인 중...")
            _classify_and_check(df, folder)
            _send_table()

            if folder:
                out = str(Path(folder) / "verification_final.xlsx")
                df.to_excel(out, index=False, engine="openpyxl")

            complete_cnt = (df["비고"] == "증빙 완비").sum() if "비고" in df.columns else 0
            missing_cnt = total - complete_cnt
            _send_progress(total, total, "전체 검증 완료!")
            log.info(f"전체 검증 완료! O:{ok}  X:{ng}  ?:{unk}  증빙완비:{complete_cnt}  확인필요:{missing_cnt}")

    except Exception as e:
        log.error(f"진위 확인 오류: {e}")
        _broadcast("error", str(e))
    finally:
        _state["running"] = False
        _broadcast("status", "idle")


# ── 2단계: 진위 확인 ──────────────────────────
@app.route("/api/verify", methods=["POST"])
def api_verify():
    if _state["running"]:
        return jsonify({"error": "이미 작업 진행 중"}), 409
    if _state["df"] is None or _state["df"].empty:
        return jsonify({"error": "먼저 QR 추출을 실행하세요"}), 400

    threading.Thread(target=_worker_verify, daemon=True).start()
    return jsonify({"ok": True})


def _worker_verify():
    import urllib3
    urllib3.disable_warnings()

    _state["running"] = True
    _broadcast("status", "running")
    try:
        verify = _import_verify()
        df = _state["df"]

        if "진위여부" not in df.columns:
            df["진위여부"] = ""

        total = len(df)
        for idx, row in df.iterrows():
            i = idx + 1
            existing = str(row.get("진위여부", "")).strip()

            if existing in ("O", "X"):
                _send_progress(i, total, f"진위 확인 중... {i}/{total}")
                continue

            if str(row.get("승인번호", "")).startswith("인식"):
                df.at[idx, "진위여부"] = "확인불가"
                _send_progress(i, total, f"진위 확인 중... {i}/{total}")
                _send_table()
                continue

            log.info(f"[{i}/{total}] {row['파일명']}")
            _send_progress(i - 1, total, f"진위 확인 중... {i}/{total}")

            row_data = {
                "승인번호_원본": str(row.get("승인번호", "")),
                "승인번호_정제": verify.clean_approval_number(str(row.get("승인번호", ""))),
                "공급자번호":    verify.clean_number(row.get("공급자번호", "")),
                "작성일자_raw":  verify.clean_date(str(row.get("작성일자", ""))),
                "합계금액":      verify.clean_number(row.get("합계금액", "")),
                "수급자번호":    "",
            }

            result = verify.verify_via_api(row_data) or "확인불가"
            df.at[idx, "진위여부"] = result
            log.info(f"  판정: {result}")
            _send_table()

            time.sleep(random.uniform(2, 5))

        _send_progress(total, total, "진위 확인 완료!")

        folder = _state["folder"]
        if folder:
            out = str(Path(folder) / "verification_final.xlsx")
            df.to_excel(out, index=False, engine="openpyxl")
            log.info(f"저장 → {out}")

        ok  = (df["진위여부"] == "O").sum()
        ng  = (df["진위여부"] == "X").sum()
        unk = (df["진위여부"] == "확인불가").sum()
        log.info(f"진위 확인 완료! O:{ok}  X:{ng}  ?:{unk}")

    except Exception as e:
        log.error(f"진위 확인 오류: {e}")
        _broadcast("error", str(e))
    finally:
        _state["running"] = False
        _broadcast("status", "idle")


# ── 중지 ─────────────────────────────────────
@app.route("/api/stop", methods=["POST"])
def api_stop():
    if not _state["running"]:
        return jsonify({"error": "실행 중인 작업 없음"}), 400
    _state["stop_requested"] = True
    log.info("중지 요청됨 — 현재 건 완료 후 중지합니다...")
    return jsonify({"ok": True})


# ── 이어서 처리 (Resume) ────────────────────
@app.route("/api/resume", methods=["POST"])
def api_resume():
    if _state["running"]:
        return jsonify({"error": "이미 작업 진행 중"}), 409

    folder = request.json.get("folder", "").strip()
    if not folder or not Path(folder).is_dir():
        return jsonify({"error": "유효한 폴더 경로를 입력하세요"}), 400

    # 기존 결과 파일 로드
    final_path = Path(folder) / "verification_final.xlsx"
    result_path = Path(folder) / "results.xlsx"

    if final_path.exists():
        _state["df"] = pd.read_excel(str(final_path))
        log.info(f"기존 결과 로드: {final_path.name}")
    elif result_path.exists():
        _state["df"] = pd.read_excel(str(result_path))
        log.info(f"기존 결과 로드: {result_path.name}")
    else:
        return jsonify({"error": "이어서 처리할 기존 결과 파일이 없습니다 (results.xlsx 또는 verification_final.xlsx)"}), 400

    _state["folder"] = folder
    _send_table()

    remaining = 0
    if "진위여부" in _state["df"].columns:
        remaining = len(_state["df"][~_state["df"]["진위여부"].isin(["O", "X", "확인불가"])])
    else:
        remaining = len(_state["df"])

    if remaining == 0:
        return jsonify({"ok": True, "remaining": 0, "message": "모든 건이 이미 확인 완료됨"})

    log.info(f"미확인 {remaining}건 이어서 처리 시작")
    _state["stop_requested"] = False
    threading.Thread(target=_resume_worker, args=(folder,), daemon=True).start()
    return jsonify({"ok": True, "remaining": remaining})


def _resume_worker(folder: str):
    _state["running"] = True
    _broadcast("status", "running")
    _run_verify_stage(folder)


# ── Excel 다운로드 ────────────────────────────
@app.route("/api/download")
def api_download():
    if _state["df"] is None or _state["df"].empty:
        return jsonify({"error": "다운로드할 데이터 없음"}), 400

    buf = io.BytesIO()
    _state["df"].to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(buf, download_name=f"verification_{ts}.xlsx",
                     as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Excel 업로드 ──────────────────────────────
@app.route("/api/upload", methods=["POST"])
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        return jsonify({"error": ".xlsx 파일을 선택하세요"}), 400
    try:
        _state["df"] = pd.read_excel(f)
        _send_table()
        return jsonify({"ok": True, "count": len(_state["df"])})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


# ── 증빙자료 조회 ────────────────────────────
@app.route("/api/evidence")
def api_evidence():
    """사용용도별 필요 증빙자료 전체 목록 반환"""
    return jsonify(_evidence_data)


@app.route("/api/evidence/<path:category>")
def api_evidence_detail(category):
    """특정 사용용도의 필요 증빙자료 반환"""
    info = _evidence_data.get(category)
    if not info:
        return jsonify({"error": "해당 사용용도가 없습니다"}), 404
    return jsonify(info)


# ── 초기화 (새로고침) ─────────────────────────
@app.route("/api/reset", methods=["POST"])
def api_reset():
    # 작업 중이면 강제 중지 후 초기화
    if _state["running"]:
        _state["stop_requested"] = True
        log.info("초기화 요청 — 작업 강제 중지 중...")
        # 워커 스레드가 중지될 때까지 최대 10초 대기
        for _ in range(20):
            time.sleep(0.5)
            if not _state["running"]:
                break
        _state["running"] = False  # 강제로 idle 전환
    _state["df"] = None
    _state["folder"] = ""
    _state["progress"] = 0
    _state["total"] = 0
    _state["message"] = "대기 중"
    _state["stop_requested"] = False
    _broadcast("table", [])
    _broadcast("progress", {"current": 0, "total": 0, "message": "대기 중"})
    _broadcast("status", "idle")
    log.info("초기화 완료 — 새 폴더를 선택하세요.")
    return jsonify({"ok": True})


# ── 폴더 선택 다이얼로그 (tkinter) ───────────
@app.route("/api/browse", methods=["POST"])
def api_browse():
    """로컬 OS 폴더 선택 다이얼로그를 열어 경로 반환"""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="PDF 폴더 선택")
    root.destroy()
    if folder:
        return jsonify({"folder": folder})
    return jsonify({"folder": ""})


# ══════════════════════════════════════════════
# HTML 프론트엔드 (Single Page)
# ══════════════════════════════════════════════
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>산업안전보건관리비 진위여부 확인 시스템</title>
<style>
:root{--primary:#1a237e;--accent:#283593;--ok:#4caf50;--fail:#e53935;--warn:#ff9800;--bg:#f5f5f5;--card:#fff;--border:#e0e0e0;--text:#212121;--text2:#757575}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'맑은 고딕','Malgun Gothic','Segoe UI',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}

/* Header */
.header{background:var(--primary);color:#fff;padding:16px 24px;display:flex;align-items:center;gap:12px;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.header h1{font-size:20px;font-weight:700;letter-spacing:-.5px}
.header .badge{background:rgba(255,255,255,.15);padding:4px 10px;border-radius:12px;font-size:11px}

/* Container */
.container{max-width:1100px;margin:0 auto;padding:20px}

/* Cards */
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:20px;margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.card-title{font-size:14px;font-weight:700;color:var(--primary);margin-bottom:12px;display:flex;align-items:center;gap:6px}

/* Folder input */
.folder-row{display:flex;gap:8px;align-items:center}
.folder-row input[type=text]{flex:1;padding:10px 14px;border:1px solid var(--border);border-radius:6px;font-size:14px;outline:none;transition:border .2s}
.folder-row input[type=text]:focus{border-color:var(--primary)}

/* Buttons */
.btn{padding:10px 20px;border:none;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;gap:6px}
.btn:disabled{opacity:.5;cursor:not-allowed}
.btn-primary{background:var(--primary);color:#fff}.btn-primary:hover:not(:disabled){background:var(--accent)}
.btn-ok{background:var(--ok);color:#fff}.btn-ok:hover:not(:disabled){background:#43a047}
.btn-outline{background:#fff;color:var(--primary);border:1px solid var(--primary)}.btn-outline:hover:not(:disabled){background:#e8eaf6}
.btn-sm{padding:7px 14px;font-size:12px}
.btn-full{background:linear-gradient(135deg,#1a237e,#42a5f5);color:#fff;font-size:14px;padding:12px 24px}.btn-full:hover:not(:disabled){background:linear-gradient(135deg,#283593,#1e88e5)}
.btn-danger{background:#e53935;color:#fff}.btn-danger:hover:not(:disabled){background:#c62828}
.btn-resume{background:#ff9800;color:#fff}.btn-resume:hover:not(:disabled){background:#f57c00}
.btn-row{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px;align-items:center}
.hint{font-size:11px;color:var(--text2);margin-top:6px}

/* Stats */
.stats{display:flex;gap:12px;flex-wrap:wrap;margin-top:12px}
.stat-chip{padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600;background:#e8eaf6;color:var(--primary)}
.stat-chip.ok{background:#e8f5e9;color:#2e7d32}
.stat-chip.fail{background:#ffebee;color:#c62828}
.stat-chip.warn{background:#fff8e1;color:#e65100}

/* Progress */
.progress-wrap{margin-top:12px;display:none}
.progress-bar-bg{height:8px;background:#e0e0e0;border-radius:4px;overflow:hidden}
.progress-bar{height:100%;background:linear-gradient(90deg,var(--primary),#42a5f5);border-radius:4px;transition:width .3s;width:0%}
.progress-text{font-size:12px;color:var(--text2);margin-top:4px}

/* Table */
.table-wrap{overflow-x:auto;margin-top:8px}
table{width:100%;border-collapse:collapse;font-size:13px}
thead{background:var(--primary);color:#fff}
th{padding:10px 12px;text-align:left;font-weight:600;white-space:nowrap}
td{padding:9px 12px;border-bottom:1px solid var(--border);white-space:nowrap}
tbody tr:hover{background:#e8eaf6}
tr.row-ok{background:#e8f5e9!important}
tr.row-fail{background:#ffebee!important}
tr.row-unk{background:#fff8e1!important}
.verdict-o{color:#2e7d32;font-weight:700}
.verdict-x{color:#c62828;font-weight:700}
.verdict-q{color:#e65100;font-weight:600}
.empty-msg{text-align:center;padding:40px;color:var(--text2);font-size:14px}
td.amt{text-align:right;font-variant-numeric:tabular-nums}
.remark-ok{color:#2e7d32;font-weight:600;font-size:12px}
.remark-warn{color:#e65100;font-size:11px}

/* Evidence */
.ev-select{padding:8px 12px;border:1px solid var(--border);border-radius:6px;font-size:13px;outline:none;min-width:200px}
.ev-select:focus{border-color:var(--primary)}
.ev-result{margin-top:12px;display:none}
.ev-result.show{display:block}
.ev-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-top:8px}
.ev-card{background:#f8f9fa;border:1px solid var(--border);border-radius:8px;padding:14px}
.ev-card-title{font-size:12px;font-weight:700;color:var(--primary);margin-bottom:8px}
.ev-card ul{list-style:none;padding:0;margin:0}
.ev-card li{font-size:12px;color:var(--text);padding:3px 0;padding-left:14px;position:relative}
.ev-card li::before{content:"";position:absolute;left:0;top:9px;width:6px;height:6px;background:var(--primary);border-radius:50%}
.ev-alt{color:var(--text2);font-size:11px;font-style:italic;margin-top:2px}

/* Log */
.log-toggle{cursor:pointer;user-select:none}
.log-toggle .arrow{display:inline-block;transition:transform .2s;font-size:12px;margin-left:6px;color:var(--text2)}
.log-toggle.open .arrow{transform:rotate(90deg)}
.log-box{background:#1e1e1e;color:#d4d4d4;font-family:Consolas,'Courier New',monospace;font-size:12px;height:160px;overflow-y:auto;padding:10px;border-radius:6px;line-height:1.6;display:none}
.log-box.show{display:block}
.log-box .info{color:#9cdcfe}.log-box .warn{color:#dcdcaa}.log-box .err{color:#f48771}

/* Upload */
.upload-area{margin-top:8px}
.upload-area input[type=file]{display:none}
.upload-label{display:inline-flex;align-items:center;gap:6px;padding:7px 14px;border:1px dashed var(--border);border-radius:6px;cursor:pointer;font-size:12px;color:var(--text2);transition:border .2s}
.upload-label:hover{border-color:var(--primary);color:var(--primary)}

@media(max-width:768px){.container{padding:12px}.folder-row{flex-direction:column}.btn-row{flex-direction:column}}
</style>
</head>
<body>

<!-- Header -->
<div class="header">
  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 12l2 2 4-4"/><path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z"/></svg>
  <h1>산업안전보건관리비 진위여부 확인 시스템</h1>
  <span class="badge" id="statusBadge">대기 중</span>
</div>

<div class="container">

  <!-- 1. 폴더 & 작업 -->
  <div class="card">
    <div class="card-title">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
      작업 설정 <span style="font-weight:400;font-size:11px;color:#757575;margin-left:8px">* 1회 최대 100건까지 업로드 가능</span>
    </div>
    <div class="folder-row">
      <input type="text" id="folderInput" placeholder="PDF 파일이 있는 폴더 경로 입력  (예: C:\invoices)">
      <button class="btn btn-outline btn-sm" onclick="browseFolder()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
        폴더 선택
      </button>
    </div>
    <div class="btn-row">
      <button class="btn btn-full" id="btnFull" onclick="startFull()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>
        진위여부 검증 시작
      </button>
      <button class="btn btn-primary btn-sm" id="btnExtract" onclick="startExtract()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 7h.01M7 12h.01M7 17h.01M12 7h5M12 12h5M12 17h5"/></svg>
        1단계: QR 추출
      </button>
      <button class="btn btn-ok btn-sm" id="btnVerify" onclick="startVerify()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="10"/></svg>
        2단계: 진위 확인
      </button>
      <button class="btn btn-resume btn-sm" id="btnResume" onclick="startResume()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 4v6h6"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>
        이어서 처리
      </button>
      <button class="btn btn-danger btn-sm" id="btnStop" onclick="stopWork()" disabled>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="6" y="6" width="12" height="12" rx="1"/></svg>
        중지
      </button>
      <button class="btn btn-outline btn-sm" id="btnReset" onclick="resetData()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M23 4v6h-6"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>
        <span id="btnResetText">새로고침</span>
      </button>
      <button class="btn btn-outline btn-sm" id="btnDownload" onclick="downloadExcel()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>
        Excel 다운로드
      </button>
      <label class="upload-label" for="fileUpload">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
        기존 Excel 불러오기
      </label>
      <input type="file" id="fileUpload" accept=".xlsx" onchange="uploadExcel(this)">
    </div>

    <!-- Progress -->
    <div class="progress-wrap" id="progressWrap">
      <div class="progress-bar-bg"><div class="progress-bar" id="progressBar"></div></div>
      <div class="progress-text" id="progressText">0/0</div>
    </div>

    <!-- Stats -->
    <div class="stats" id="statsArea"></div>
  </div>

  <!-- 2. 결과 테이블 -->
  <div class="card">
    <div class="card-title">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>
      조회 결과
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr><th>#</th><th>파일명</th><th>페이지</th><th>승인번호</th><th>공급자번호</th><th>작성일자</th><th>합계금액</th><th>진위여부</th><th>사용용도</th><th>비고</th></tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="10" class="empty-msg">PDF 폴더를 지정하고 [진위여부 검증 시작]을 실행하세요</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <!-- 3. 증빙자료 안내 -->
  <div class="card">
    <div class="card-title">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
      사용용도별 필요 증빙자료 안내
    </div>
    <select class="ev-select" id="evSelect" onchange="showEvidence()">
      <option value="">-- 사용용도를 선택하세요 --</option>
    </select>
    <div class="ev-result" id="evResult">
      <div class="ev-grid" id="evGrid"></div>
    </div>
  </div>

  <!-- 4. 로그 -->
  <div class="card">
    <div class="card-title log-toggle" id="logToggle" onclick="toggleLog()">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="4 17 10 11 4 5"/><line x1="12" y1="19" x2="20" y2="19"/></svg>
      실행 로그 <span class="arrow">&#9654;</span>
    </div>
    <div class="log-box" id="logBox"></div>
  </div>

</div>

<script>
// ── SSE 연결 ──────────────────────────────────
const sse = new EventSource('/stream');

sse.addEventListener('log', e => {
  const box = document.getElementById('logBox');
  const line = document.createElement('div');
  const text = e.data;
  if (text.includes('[ERROR]'))      line.className = 'err';
  else if (text.includes('[WARNING]')) line.className = 'warn';
  else                                line.className = 'info';
  line.textContent = text;
  box.appendChild(line);
  box.scrollTop = box.scrollHeight;
});

sse.addEventListener('progress', e => {
  const d = JSON.parse(e.data);
  const wrap = document.getElementById('progressWrap');
  wrap.style.display = 'block';
  const pct = d.total > 0 ? Math.round(d.current / d.total * 100) : 0;
  document.getElementById('progressBar').style.width = pct + '%';
  document.getElementById('progressText').textContent = d.message || `${d.current}/${d.total}`;
});

sse.addEventListener('table', e => {
  const rows = JSON.parse(e.data);
  renderTable(rows);
  fetchStatus();
});

sse.addEventListener('status', e => {
  const s = e.data.replace(/"/g, '');
  const badge = document.getElementById('statusBadge');
  const running = s === 'running';
  badge.textContent = running ? '처리 중...' : '대기 중';
  badge.style.background = running ? 'rgba(255,255,255,.3)' : 'rgba(255,255,255,.15)';
  document.getElementById('btnFull').disabled = running;
  document.getElementById('btnExtract').disabled = running;
  document.getElementById('btnVerify').disabled = running;
  document.getElementById('btnResume').disabled = running;
  document.getElementById('btnStop').disabled = !running;
});

sse.addEventListener('error', e => {
  alert('오류: ' + e.data);
});

// ── 테이블 렌더링 ─────────────────────────────
function renderTable(rows) {
  const tbody = document.getElementById('tableBody');
  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" class="empty-msg">데이터 없음</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => {
    let cls = '';
    let vcls = '';
    if (r.verdict === 'O')       { cls = 'row-ok';   vcls = 'verdict-o'; }
    else if (r.verdict === 'X')  { cls = 'row-fail'; vcls = 'verdict-x'; }
    else if (r.verdict === '확인불가') { cls = 'row-unk'; vcls = 'verdict-q'; }
    const remark = r.remark || '';
    let rcls = '';
    if (remark === '증빙 완비') rcls = 'remark-ok';
    else if (remark.includes('확인 필요')) rcls = 'remark-warn';
    return `<tr class="${cls}">
      <td>${r.no}</td><td>${esc(r.filename)}</td><td>${r.page}</td>
      <td>${esc(r.approval)}</td><td>${r.supplier}</td><td>${r.date}</td>
      <td class="amt">${r.amount}</td><td class="${vcls}">${r.verdict || '-'}</td>
      <td>${esc(r.usage || '')}</td><td class="${rcls}">${esc(remark) || '-'}</td>
    </tr>`;
  }).join('');
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

// ── 통계 ──────────────────────────────────────
function fetchStatus() {
  fetch('/api/status').then(r => r.json()).then(d => {
    const area = document.getElementById('statsArea');
    const s = d.stats;
    if (!s || !s.total) { area.innerHTML = ''; return; }
    area.innerHTML = `
      <span class="stat-chip">전체 ${s.total}건</span>
      <span class="stat-chip">QR인식 ${s.qr_ok}건</span>
      <span class="stat-chip ok">O 정상 ${s.ok}건</span>
      <span class="stat-chip fail">X 불일치 ${s.fail}건</span>
      <span class="stat-chip warn">? 확인불가 ${s.unknown}건</span>
    `;
  });
}

// ── 버튼 핸들러 ───────────────────────────────
function browseFolder() {
  fetch('/api/browse', { method: 'POST' })
    .then(r => r.json())
    .then(d => { if (d.folder) document.getElementById('folderInput').value = d.folder; });
}

function startFull() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) { alert('PDF 폴더 경로를 입력하세요.'); return; }
  fetch('/api/full', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({folder})
  }).then(r => r.json()).then(d => {
    if (d.error) alert(d.error);
  });
}

function startExtract() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) { alert('PDF 폴더 경로를 입력하세요.'); return; }
  fetch('/api/extract', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({folder})
  }).then(r => r.json()).then(d => {
    if (d.error) alert(d.error);
  });
}

function stopWork() {
  fetch('/api/stop', { method: 'POST' })
    .then(r => r.json())
    .then(d => { if (d.error) alert(d.error); });
}

function startResume() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) { alert('PDF 폴더 경로를 입력하세요.'); return; }
  fetch('/api/resume', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({folder})
  }).then(r => r.json()).then(d => {
    if (d.error) alert(d.error);
    else if (d.remaining === 0) alert(d.message);
  });
}

function startVerify() {
  fetch('/api/verify', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
  }).then(r => r.json()).then(d => {
    if (d.error) alert(d.error);
  });
}

function resetData() {
  if (!confirm('현재 결과를 초기화하고 새 폴더를 선택하시겠습니까?')) return;
  const btn = document.getElementById('btnReset');
  const txt = document.getElementById('btnResetText');
  btn.disabled = true;
  txt.textContent = '초기화 중...';
  fetch('/api/reset', { method: 'POST' })
    .then(r => r.json())
    .then(d => {
      if (d.error) { alert(d.error); return; }
      document.getElementById('folderInput').value = '';
      document.getElementById('statsArea').innerHTML = '';
      document.getElementById('progressWrap').style.display = 'none';
      document.getElementById('progressBar').style.width = '0%';
      document.getElementById('progressText').textContent = '';
      document.getElementById('tableBody').innerHTML = '<tr><td colspan="10" class="empty-msg">PDF 폴더를 지정하고 [진위여부 검증 시작]을 실행하세요</td></tr>';
    })
    .finally(() => {
      btn.disabled = false;
      txt.textContent = '새로고침';
    });
}

function downloadExcel() {
  window.location.href = '/api/download';
}

function uploadExcel(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  fetch('/api/upload', { method: 'POST', body: fd })
    .then(r => r.json())
    .then(d => {
      if (d.error) alert(d.error);
      else fetchStatus();
    });
  input.value = '';
}

// ── 증빙자료 안내 ────────────────────────────
let _evData = {};

function loadEvidence() {
  fetch('/api/evidence').then(r => r.json()).then(data => {
    _evData = data;
    const sel = document.getElementById('evSelect');
    Object.keys(data).forEach(cat => {
      const opt = document.createElement('option');
      opt.value = cat;
      opt.textContent = cat;
      sel.appendChild(opt);
    });
  });
}

function showEvidence() {
  const cat = document.getElementById('evSelect').value;
  const wrap = document.getElementById('evResult');
  const grid = document.getElementById('evGrid');
  if (!cat || !_evData[cat]) { wrap.classList.remove('show'); return; }

  const info = _evData[cat];
  const labels = {'증빙자료1': '증빙자료 1', '증빙자료2': '증빙자료 2', '증빙자료3': '증빙자료 3'};
  let html = '';

  for (const [key, items] of Object.entries(info)) {
    if (!items || items.length === 0) continue;
    html += `<div class="ev-card"><div class="ev-card-title">${labels[key] || key}</div><ul>`;
    items.forEach((item, idx) => {
      if (idx > 0) html += `<div class="ev-alt">또는</div>`;
      html += `<li>${esc(item)}</li>`;
    });
    html += '</ul></div>';
  }

  grid.innerHTML = html || '<p style="color:#757575;font-size:13px">등록된 증빙자료 정보가 없습니다.</p>';
  wrap.classList.add('show');
}

// ── 로그 토글 ──────────────────────────────────
function toggleLog() {
  document.getElementById('logToggle').classList.toggle('open');
  document.getElementById('logBox').classList.toggle('show');
}

// 초기 로드
fetchStatus();
loadEvidence();
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


# ══════════════════════════════════════════════
# 진입점
# ══════════════════════════════════════════════
if __name__ == "__main__":
    import webbrowser, urllib3
    urllib3.disable_warnings()

    port = 5000
    print("\n" + "="*55)
    print("  산업안전보건관리비 진위여부 확인 시스템 (웹)")
    print("="*55)
    print(f"  브라우저에서 접속: http://localhost:{port}")
    print("  종료: Ctrl+C")
    print("="*55 + "\n")

    webbrowser.open(f"http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
