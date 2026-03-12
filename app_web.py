"""
산업안전보건관리비 진위여부 확인 시스템 (웹 UI)
================================================
Flask 기반 웹 애플리케이션 — 브라우저에서 접속하여 사용

실행: python app_web.py
접속: http://localhost:5000

필수: pip install flask pandas openpyxl PyMuPDF pyzbar opencv-python-headless
              pillow zxing-cpp requests google-generativeai
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
    "gemini_api_key": "",  # Gemini API Key
    "list_df": None,       # 집행실적리스트 DataFrame
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
    is_ai = "매칭PDF" in _state["df"].columns  # AI 모드 여부

    for i, (_, r) in enumerate(_state["df"].iterrows()):
        if is_ai:
            amt = r.get("공급가액", "")
            try:
                amt = f"{int(str(amt).replace(',', '')):,}"
            except (ValueError, TypeError):
                amt = str(amt)
            rows.append({
                "mode": "ai",
                "no": i + 1,
                "usage": str(r.get("사용용도", "")),
                "item": str(r.get("품목", "")),
                "doc_type": str(r.get("전표구분", "")),
                "approval": str(r.get("승인번호", "")),
                "supplier_name": str(r.get("공급사업자명", "")),
                "amount": amt,
                "matched_pdf": str(r.get("매칭PDF", "")),
                "hometax": str(r.get("홈택스진위", "")),
                "evidence_ok": str(r.get("증빙적합", "")),
                "ai_result": str(r.get("AI검증", "")),
                "remark": str(r.get("비고", "")),
            })
        else:
            amt = r.get("합계금액", "")
            try:
                amt = f"{int(str(amt).replace(',', '')):,}"
            except (ValueError, TypeError):
                amt = str(amt)
            rows.append({
                "mode": "qr",
                "no": i + 1,
                "filename": str(r.get("파일명", "")),
                "page": str(r.get("인식페이지", "")),
                "approval": str(r.get("승인번호", "")),
                "supplier": str(r.get("공급자번호", "")),
                "date": str(r.get("작성일자", "")),
                "amount": amt,
                "verdict": str(r.get("진위여부", "")),
                "usage": str(r.get("사용용도", "")),
                "suitability": str(r.get("적합성", "")),
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
    """data.xlsx → 사용용도별 증빙자료 구조화 (없으면 내장 데이터 사용)"""
    data_path = Path(__file__).parent / "data.xlsx"
    if data_path.exists():
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
            extra = str(row.get("Unnamed: 5", "")).strip()
            if extra and extra != "nan" and extra not in result[key].get("증빙자료3", []):
                result[key]["증빙자료3"].append(extra)
        return result

    # data.xlsx가 없으면 내장 데이터 사용
    _BUILTIN = {
        "1.안전관리자 임금 등 (안전관리자)": {
            "증빙자료1": ["안전관리자 선임 신고서"],
            "증빙자료2": ["급여명세서"],
            "증빙자료3": [],
        },
        "1.안전관리자 임금 등 (보건관리자)": {
            "증빙자료1": ["보건관리자 선임 신고서"],
            "증빙자료2": ["급여명세서"],
            "증빙자료3": [],
        },
        "1.안전관리자 임금 등 (신호수, 유도원)": {
            "증빙자료1": ["근로계약서"],
            "증빙자료2": ["지급내역서"],
            "증빙자료3": [],
        },
        "2.안전시설비": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서"],
            "증빙자료2": ["거래명세표"],
            "증빙자료3": [],
        },
        "3.보호구": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서"],
            "증빙자료2": ["거래명세표"],
            "증빙자료3": ["지급대장(안전장구 착용 사진 포함)"],
        },
        "4.안전보건진단비": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서"],
            "증빙자료2": ["계약서"],
            "증빙자료3": ["유자격확인자료"],
        },
        "5.안전보건교육비": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서", "사업자 현금영수증"],
            "증빙자료2": ["수료증", "거래명세표"],
            "증빙자료3": [],
        },
        "6.근로자 건강장해예방비": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서", "사업자 현금영수증"],
            "증빙자료2": ["입증자료(구입사진 등)"],
            "증빙자료3": [],
        },
        "7.기술지도비": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서"],
            "증빙자료2": ["계약서"],
            "증빙자료3": ["결과보고서"],
        },
        "8.본사 전담조직근로자 임금": {
            "증빙자료1": ["안전전담 증빙자료"],
            "증빙자료2": ["급여명세서"],
            "증빙자료3": [],
        },
        "9.위험성평가 등에 따른 소요비용": {
            "증빙자료1": ["전자세금계산서", "법인카드전표 또는 법인카드 이용내역서"],
            "증빙자료2": ["노사협의체 협의록"],
            "증빙자료3": ["위험성평가표", "위험성평가 결과 검토서(발주부서 승인분)"],
        },
    }
    return _BUILTIN


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


# ──────────────────────────────────────────────
# Gemini Vision AI 파이프라인
# ──────────────────────────────────────────────
def _init_gemini():
    """Gemini 모델 초기화 (api key 필요)"""
    import google.generativeai as genai
    genai.configure(api_key=_state["gemini_api_key"])
    return genai.GenerativeModel("gemini-1.5-flash")



def _parse_execution_list(file_path):
    """집행실적리스트 엑셀 파싱 — 6행 헤더, 7행부터 데이터"""
    df = pd.read_excel(file_path, engine="openpyxl", header=None, skiprows=7)
    col_map = {
        0: "순번", 1: "계약번호", 2: "계약명", 3: "기성준공", 4: "회차",
        5: "승인번호", 6: "발행일자", 7: "사용용도", 8: "품목", 9: "품목기타",
        10: "수량", 11: "단위금액", 12: "공급가액", 13: "인정금액",
        14: "전표구분", 15: "국세청조회", 16: "공급사업자번호", 17: "공급사업자명",
        18: "진행단계",
    }
    df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)
    # 순번이 비어있는 행 제거
    df = df.dropna(subset=["순번"]).reset_index(drop=True)
    # 승인번호를 문자열로
    df["승인번호"] = df["승인번호"].astype(str).str.strip()
    return df


def _gemini_extract_approval(model, pdf_path):
    """Gemini Vision으로 PDF의 모든 페이지에서 24자리 승인번호를 OCR 추출 (복수 지원)"""
    import fitz
    from PIL import Image
    import re as _re

    doc = fitz.open(str(pdf_path))
    images = []
    for page_num in range(min(len(doc), 10)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=200)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()

    prompt = """이 세금계산서/영수증 이미지들에서 국세청 승인번호(24자리 숫자+영문)를 모두 찾아줘.
한 페이지에 세금계산서가 여러 장 있을 수 있으니 모든 승인번호를 빠짐없이 추출해.
승인번호는 보통 "승인번호", "국세청승인번호" 옆에 있거나,
8자리-8자리-8자리 형태(예: 20240306-12345678-90123456)로 표기돼.

[응답 포맷] (순수 JSON만 출력, 마크다운 없이)
{ "approval_numbers": ["승인번호1", "승인번호2", ...] }
승인번호가 없으면 빈 배열로 응답해."""

    try:
        content = [prompt] + images
        response = model.generate_content(content)
        text = response.text.replace("```json", "").replace("```", "").strip()
        result = json.loads(text)

        approvals = []
        for raw in result.get("approval_numbers", []):
            clean = _re.sub(r'[^0-9a-zA-Z]', '', str(raw))
            if len(clean) == 24:
                approvals.append(clean)
        return approvals if approvals else None
    except Exception as e:
        log.warning(f"Gemini OCR 승인번호 추출 실패 ({pdf_path.name}): {e}")
        return None


def _gemini_verify_evidence(model, pdf_path, excel_row):
    """Gemini Vision으로 PDF 내용과 엑셀 데이터 대조 검증
    - 품목/사용용도가 PDF 내용과 부합하는지
    - 필요 증빙서류가 PDF에 포함되어 있는지
    한 번의 API 호출로 처리"""
    import fitz
    from PIL import Image

    doc = fitz.open(str(pdf_path))
    images = []
    for page_num in range(min(len(doc), 10)):  # 최대 10페이지
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()

    usage = str(excel_row.get("사용용도", ""))
    item = str(excel_row.get("품목", ""))
    item_etc = str(excel_row.get("품목기타", ""))
    doc_type = str(excel_row.get("전표구분", ""))
    amount = str(excel_row.get("공급가액", ""))
    supplier = str(excel_row.get("공급사업자명", ""))

    # 사용용도에서 번호 제거 (예: "2.안전시설비 등" → "안전시설비")
    usage_clean = usage.split(".")[-1].replace(" 등", "").strip() if "." in usage else usage

    # 필요 증빙 목록
    needed = []
    for cat, evidence_list in _ADDITIONAL_EVIDENCE.items():
        if cat in usage_clean or usage_clean in cat:
            needed = evidence_list
            break
    needed_str = ", ".join(needed) if needed else "추가 증빙 없음"

    prompt = f"""너는 산업안전보건관리비 감사관이야. 아래 PDF 이미지들이 집행실적 데이터와 부합하는지 검증해.

[집행실적 데이터]
- 사용용도: {usage}
- 품목: {item}
- 품목(기타): {item_etc}
- 전표구분: {doc_type}
- 공급가액: {amount}
- 공급사업자: {supplier}

[필요 증빙서류 목록]
이 사용용도({usage})에는 전표(세금계산서/카드전표 등) 외에 다음 추가 증빙서류가 반드시 필요해:
{needed_str}

[검증 항목]
1. 품목 부합: PDF 내용(세금계산서/영수증의 품목명)이 위 품목과 일치하거나 관련 있는지
2. 증빙서류 적합: 위 필요 증빙서류 목록의 서류들이 이 PDF 파일에 실제로 포함되어 있는지 각각 확인
3. 금액 확인: PDF의 금액이 집행실적 데이터와 일치하는지

[응답 포맷] (순수 JSON만 출력, 마크다운 없이)
{{ "match": "PASS" 또는 "FAIL" 또는 "REVIEW", "reason": "검증 결과 요약 1~2문장", "evidence_ok": "O" 또는 "X", "found_docs": ["PDF에서 실제 감지된 증빙서류들"], "missing_docs": ["누락된 필요 증빙서류들"] }}

증빙서류 적합 판단 기준:
- 필요 증빙서류가 모두 PDF에 포함되어 있으면 evidence_ok = "O"
- 하나라도 누락되면 evidence_ok = "X"
- 필요 증빙서류가 없는 경우(추가 증빙 없음) evidence_ok = "O"
"""

    try:
        content = [prompt] + images
        response = model.generate_content(content)
        text = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)
    except Exception as e:
        return {"match": "REVIEW", "reason": f"AI 검증 오류: {e}", "found_docs": [], "missing_docs": []}


def _worker_full_ai(folder: str, pdf_files: list):
    """AI 모드: 엑셀 기반 매칭 → 홈택스 진위 → AI 증빙 검증"""
    import urllib3
    urllib3.disable_warnings()

    _state["running"] = True
    _state["stop_requested"] = False
    _broadcast("status", "running")

    try:
        model = _init_gemini()
        verify = _import_verify()
        qr = _import_qr()

        # ── 1단계: 엑셀 로드 + PDF QR 추출 + 매칭 ──
        list_df = _state.get("list_df")
        if list_df is None or list_df.empty:
            _broadcast("error", "집행실적 엑셀이 업로드되지 않았습니다.")
            return

        total = len(list_df)
        log.info(f"집행실적 {total}건, PDF {len(pdf_files)}개 — 매칭 시작")

        # PDF QR 추출 (멀티 QR 지원) + Gemini OCR 보완
        qr_map = {}  # 승인번호 → {pdf_path, qr_data}
        ocr_fallback_count = 0
        # 엑셀에 있는 승인번호 목록 (매칭 대상)
        excel_approvals = set(list_df["승인번호"].astype(str).str.strip().tolist())

        for i, pdf_path in enumerate(pdf_files):
            if _state["stop_requested"]:
                break
            _send_progress(i, len(pdf_files), f"1단계: PDF QR 추출 중... {i+1}/{len(pdf_files)}")
            records = qr.process_pdf_multi(str(pdf_path))
            qr_found = 0
            for record in records:
                approval = str(record.get("승인번호", "")).strip()
                if approval and approval != "인식 불가":
                    qr_map[approval] = {"path": pdf_path, "record": record}
                    qr_found += 1

            # 엑셀에서 이 PDF에 매칭 안 된 승인번호가 아직 있으면 Gemini OCR 시도
            unmatched = excel_approvals - set(qr_map.keys())
            if unmatched and (qr_found == 0 or len(unmatched) > 0):
                _send_progress(i, len(pdf_files), f"1단계: AI OCR 보완 중... {pdf_path.name}")
                ocr_approvals = _gemini_extract_approval(model, pdf_path)
                if ocr_approvals:
                    for ocr_app in ocr_approvals:
                        if ocr_app not in qr_map:
                            qr_map[ocr_app] = {"path": pdf_path, "record": {"승인번호": ocr_app, "OCR": True}}
                            ocr_fallback_count += 1
                            log.info(f"OCR 보완 성공: {pdf_path.name} → {ocr_app}")
                elif qr_found == 0:
                    log.warning(f"QR+OCR 모두 실패: {pdf_path.name}")

        log.info(f"QR 추출 완료: {len(qr_map)}개 승인번호 인식 (OCR 보완: {ocr_fallback_count}건)")

        # 엑셀 ↔ PDF 매칭 + 결과 DataFrame 구성
        records = []
        for idx, row in list_df.iterrows():
            excel_approval = str(row.get("승인번호", "")).strip()
            matched_pdf = ""
            qr_data = None

            # 승인번호로 매칭
            if excel_approval in qr_map:
                matched_pdf = qr_map[excel_approval]["path"].name
                qr_data = qr_map[excel_approval]["record"]

            records.append({
                "순번": row.get("순번", ""),
                "사용용도": row.get("사용용도", ""),
                "품목": row.get("품목", ""),
                "전표구분": row.get("전표구분", ""),
                "승인번호": excel_approval,
                "공급사업자명": row.get("공급사업자명", ""),
                "공급가액": row.get("공급가액", ""),
                "매칭PDF": matched_pdf,
                "홈택스진위": "",
                "증빙적합": "",
                "AI검증": "",
                "비고": "증빙없음" if not matched_pdf else "",
            })

        col_order = ["순번", "사용용도", "품목", "전표구분", "승인번호",
                      "공급사업자명", "공급가액", "매칭PDF", "홈택스진위", "증빙적합", "AI검증", "비고"]
        _state["df"] = pd.DataFrame(records, columns=col_order)
        matched_cnt = (_state["df"]["매칭PDF"] != "").sum()
        log.info(f"1단계 완료: 매칭 {matched_cnt}/{total}")
        _send_progress(total, total, f"1단계 완료! 매칭 {matched_cnt}/{total}")
        _send_table()

        if _state["stop_requested"]:
            _state["running"] = False
            _broadcast("status", "idle")
            return

        # ── 2단계: 홈택스 진위확인 (전자세금계산서만, 동일 승인번호 중복 건너뛰기) ──
        df = _state["df"]
        hometax_cache = {}  # 승인번호 → 결과 (중복 방지)

        for idx, row in df.iterrows():
            if _state["stop_requested"]:
                break
            i = idx + 1
            doc_type = str(row.get("전표구분", ""))
            matched_pdf = str(row.get("매칭PDF", ""))
            excel_approval = str(row.get("승인번호", ""))

            if "세금계산서" not in doc_type:
                df.at[idx, "홈택스진위"] = "-"
                _send_progress(i, total, f"2단계: 진위 확인 중... {i}/{total}")
                _send_table()
                continue

            if not matched_pdf:
                df.at[idx, "홈택스진위"] = "증빙없음"
                _send_progress(i, total, f"2단계: 진위 확인 중... {i}/{total}")
                _send_table()
                continue

            # 동일 승인번호 이미 확인했으면 결과 재사용
            if excel_approval in hometax_cache:
                df.at[idx, "홈택스진위"] = hometax_cache[excel_approval]
                log.info(f"[2단계 {i}/{total}] 중복 건너뛰기: {row['품목']} → {hometax_cache[excel_approval]}")
                _send_progress(i, total, f"2단계: 진위 확인 중... {i}/{total} (중복 스킵)")
                _send_table()
                continue

            qr_data = qr_map.get(excel_approval, {}).get("record", {})

            log.info(f"[2단계 {i}/{total}] 홈택스: {row['품목']}")
            _send_progress(i - 1, total, f"2단계: 진위 확인 중... {i}/{total}")

            row_data = {
                "승인번호_원본": str(qr_data.get("승인번호", excel_approval)),
                "승인번호_정제": verify.clean_approval_number(str(qr_data.get("승인번호", excel_approval))),
                "공급자번호":    verify.clean_number(qr_data.get("공급자번호", row.get("공급사업자번호", ""))),
                "작성일자_raw":  verify.clean_date(str(qr_data.get("작성일자", row.get("발행일자", "")))),
                "합계금액":      verify.clean_number(qr_data.get("합계금액", row.get("공급가액", ""))),
                "수급자번호":    "",
            }
            result = verify.verify_via_api(row_data) or "확인불가"
            df.at[idx, "홈택스진위"] = result
            hometax_cache[excel_approval] = result
            log.info(f"  판정: {result}")
            _send_table()
            time.sleep(random.uniform(2, 5))

        _send_progress(total, total, "2단계 완료! 3단계 AI 검증 시작...")

        if _state["stop_requested"]:
            _state["running"] = False
            _broadcast("status", "idle")
            return

        # ── 3단계: Gemini AI 증빙 검증 (동일 승인번호 중복 건너뛰기) ──
        ai_cache = {}  # 승인번호 → result

        for idx, row in df.iterrows():
            if _state["stop_requested"]:
                break
            i = idx + 1
            matched_pdf = str(row.get("매칭PDF", ""))
            excel_approval = str(row.get("승인번호", ""))

            _send_progress(i - 1, total, f"3단계: AI 검증 중... {i}/{total}")

            if not matched_pdf:
                df.at[idx, "AI검증"] = "증빙없음"
                df.at[idx, "증빙적합"] = "-"
                df.at[idx, "비고"] = "PDF 증빙 없음"
                _send_table()
                continue

            pdf_path = Path(folder) / matched_pdf
            if not pdf_path.exists():
                df.at[idx, "AI검증"] = "파일없음"
                df.at[idx, "증빙적합"] = "-"
                df.at[idx, "비고"] = "PDF 파일 없음"
                _send_table()
                continue

            # 동일 승인번호 이미 검증했으면 결과 재사용
            if excel_approval in ai_cache:
                cached = ai_cache[excel_approval]
                df.at[idx, "AI검증"] = cached.get("match", "REVIEW")
                df.at[idx, "증빙적합"] = cached.get("evidence_ok", "-")
                missing = cached.get("missing_docs", [])
                if missing:
                    df.at[idx, "비고"] = f"누락: {', '.join(missing)}"
                elif cached.get("match") == "PASS":
                    df.at[idx, "비고"] = cached.get("reason", "검증 통과")
                elif cached.get("match") == "FAIL":
                    df.at[idx, "비고"] = cached.get("reason", "불부합")
                else:
                    df.at[idx, "비고"] = cached.get("reason", "확인 필요")
                log.info(f"[3단계 {i}/{total}] 중복 건너뛰기: {row['품목']} → {cached.get('match')}")
                _send_table()
                continue

            log.info(f"[3단계 {i}/{total}] AI 검증: {row['품목']}")
            result = _gemini_verify_evidence(model, pdf_path, row)
            ai_cache[excel_approval] = result

            match_status = result.get("match", "REVIEW")
            reason = result.get("reason", "")
            missing = result.get("missing_docs", [])
            evidence_ok = result.get("evidence_ok", "-")

            df.at[idx, "AI검증"] = match_status
            df.at[idx, "증빙적합"] = evidence_ok
            if missing:
                df.at[idx, "비고"] = f"누락: {', '.join(missing)}"
            elif match_status == "PASS":
                df.at[idx, "비고"] = reason or "검증 통과"
            elif match_status == "FAIL":
                df.at[idx, "비고"] = reason or "불부합"
            else:
                df.at[idx, "비고"] = reason or "확인 필요"

            log.info(f"  → {match_status} | {reason}")
            _send_table()
            time.sleep(1)

        # ── 최종 저장 ──
        if folder:
            out = str(Path(folder) / "verification_final.xlsx")
            df.to_excel(out, index=False, engine="openpyxl")
            log.info(f"저장 → {out}")

        ok = (df["홈택스진위"] == "O").sum()
        ng = (df["홈택스진위"] == "X").sum()
        ai_pass = (df["AI검증"] == "PASS").sum()
        ai_fail = (df["AI검증"] == "FAIL").sum()

        _send_progress(total, total, "AI 전체 검증 완료!")
        log.info(f"AI 전체 검증 완료! 홈택스 O:{ok} X:{ng} | AI PASS:{ai_pass} FAIL:{ai_fail}")

    except Exception as e:
        log.error(f"Gemini AI 오류: {e}")
        _broadcast("error", str(e))
    finally:
        _state["running"] = False
        _broadcast("status", "idle")


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
        is_ai = "매칭PDF" in df.columns
        if is_ai:
            stats = {
                "mode": "ai",
                "total":      len(df),
                "matched":    int((df["매칭PDF"] != "").sum()),
                "hometax_ok": int((df.get("홈택스진위") == "O").sum()) if "홈택스진위" in df.columns else 0,
                "hometax_x":  int((df.get("홈택스진위") == "X").sum()) if "홈택스진위" in df.columns else 0,
                "ai_pass":    int((df.get("AI검증") == "PASS").sum()) if "AI검증" in df.columns else 0,
                "ai_fail":    int((df.get("AI검증") == "FAIL").sum()) if "AI검증" in df.columns else 0,
                "ai_review":  int((df.get("AI검증") == "REVIEW").sum()) if "AI검증" in df.columns else 0,
                "ev_ok":      int((df.get("증빙적합") == "O").sum()) if "증빙적합" in df.columns else 0,
                "ev_x":       int((df.get("증빙적합") == "X").sum()) if "증빙적합" in df.columns else 0,
                "no_evidence":int((df.get("비고") == "증빙없음").sum()) if "비고" in df.columns else 0,
            }
        else:
            stats = {
                "mode": "qr",
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
            multi = qr.process_pdf_multi(str(pdf_path))
            records.extend(multi)

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
            multi = qr.process_pdf_multi(str(pdf_path))
            records.extend(multi)

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


# ── Gemini API Key 설정 ──────────────────────
@app.route("/api/set-apikey", methods=["POST"])
def api_set_apikey():
    key = request.json.get("key", "").strip()
    if not key:
        _state["gemini_api_key"] = ""
        return jsonify({"ok": True, "message": "API Key 초기화됨"})
    _state["gemini_api_key"] = key
    log.info("Gemini API Key 설정 완료")
    return jsonify({"ok": True, "message": "API Key 설정 완료"})


# ── 집행실적 엑셀 업로드 ─────────────────────
@app.route("/api/upload-list", methods=["POST"])
def api_upload_list():
    if "file" not in request.files:
        return jsonify({"error": "파일이 없습니다"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "엑셀 파일(.xlsx)만 업로드 가능합니다"}), 400
    try:
        import tempfile, os
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        f.save(tmp.name)
        tmp.close()
        list_df = _parse_execution_list(tmp.name)
        os.unlink(tmp.name)
        _state["list_df"] = list_df
        log.info(f"집행실적 엑셀 로드 완료: {len(list_df)}건")
        return jsonify({"ok": True, "count": len(list_df)})
    except Exception as e:
        log.error(f"엑셀 파싱 오류: {e}")
        return jsonify({"error": f"엑셀 파싱 오류: {e}"}), 400


# ── Gemini AI 모드 전체 실행 ─────────────────
@app.route("/api/full-ai", methods=["POST"])
def api_full_ai():
    if _state["running"]:
        return jsonify({"error": "이미 작업 진행 중"}), 409
    if not _state["gemini_api_key"]:
        return jsonify({"error": "Gemini API Key를 먼저 설정하세요"}), 400
    if _state.get("list_df") is None or _state["list_df"].empty:
        return jsonify({"error": "집행실적 엑셀을 먼저 업로드하세요"}), 400

    folder = request.json.get("folder", "").strip()
    if not folder or not Path(folder).is_dir():
        return jsonify({"error": "유효한 폴더 경로를 입력하세요"}), 400

    pdf_files = sorted(Path(folder).glob("*.pdf"))
    if not pdf_files:
        return jsonify({"error": "폴더에 PDF 파일이 없습니다"}), 400

    _state["folder"] = folder
    threading.Thread(target=_worker_full_ai, args=(folder, pdf_files), daemon=True).start()
    return jsonify({"ok": True, "count": len(pdf_files), "list_count": len(_state["list_df"])})


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
.suit-pass{color:#2e7d32;font-weight:700;font-size:12px}
.suit-fail{color:#c62828;font-weight:700;font-size:12px}
.suit-review{color:#e65100;font-weight:600;font-size:12px}

/* Mode selector */
.mode-row{display:flex;gap:16px;align-items:center;margin-bottom:12px;flex-wrap:wrap}
.mode-radio{display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;font-weight:600}
.mode-radio input[type=radio]{accent-color:var(--primary)}
.apikey-row{display:flex;gap:8px;align-items:center;margin-bottom:12px;display:none}
.apikey-row.show{display:flex}
.apikey-row input[type=text]{flex:1;max-width:400px;padding:8px 12px;border:1px solid var(--border);border-radius:6px;font-size:13px;outline:none;font-family:monospace}
.apikey-row input[type=text]:focus{border-color:var(--primary)}
.apikey-status{font-size:11px;font-weight:600;padding:4px 10px;border-radius:12px}
.apikey-status.set{background:#e8f5e9;color:#2e7d32}
.apikey-status.unset{background:#fff8e1;color:#e65100}
.list-upload-row{display:flex;gap:8px;align-items:center;margin-bottom:12px;display:none}
.list-upload-row.show{display:flex}
.list-upload-row input[type=file]{font-size:13px}
.list-status{font-size:11px;font-weight:600;padding:4px 10px;border-radius:12px}
.list-status.loaded{background:#e8f5e9;color:#2e7d32}
.list-status.empty{background:#fff8e1;color:#e65100}

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

    <!-- 모드 선택 -->
    <div class="mode-row">
      <label class="mode-radio"><input type="radio" name="mode" value="qr" checked onchange="onModeChange()"> 기본 모드 (QR 추출)</label>
      <label class="mode-radio"><input type="radio" name="mode" value="ai" onchange="onModeChange()"> AI 모드 (Gemini Vision)</label>
    </div>

    <!-- Gemini API Key 입력 (AI 모드 선택 시 표시) -->
    <div class="apikey-row" id="apikeyRow">
      <input type="text" id="apikeyInput" placeholder="Gemini API Key 입력">
      <button class="btn btn-outline btn-sm" onclick="setApiKey()">설정</button>
      <span class="apikey-status unset" id="apikeyStatus">미설정</span>
    </div>

    <!-- 집행실적 엑셀 업로드 (AI 모드 선택 시 표시) -->
    <div class="list-upload-row" id="listUploadRow">
      <span style="font-size:13px;font-weight:600;color:var(--primary);white-space:nowrap">산업안전보건관리비 집행 리스트</span>
      <input type="file" id="listFileInput" accept=".xlsx,.xls">
      <button class="btn btn-outline btn-sm" onclick="uploadList()">업로드</button>
      <span class="list-status empty" id="listStatus">미업로드</span>
    </div>

    <div class="folder-row">
      <input type="text" id="folderInput" placeholder="PDF 파일이 있는 폴더 경로 입력  (예: C:\invoices)">
      <button class="btn btn-outline btn-sm" onclick="browseFolder()">폴더 선택</button>
    </div>
    <div class="btn-row">
      <button class="btn btn-full" id="btnFull" onclick="startFull()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>
        <span id="btnFullText">진위여부 검증 시작</span>
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
        <thead id="tableHead">
          <tr><th>#</th><th>파일명</th><th>페이지</th><th>승인번호</th><th>공급자번호</th><th>작성일자</th><th>합계금액</th><th>진위여부</th><th>사용용도</th><th>적합성</th><th>비고</th></tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="11" class="empty-msg">PDF 폴더를 지정하고 [진위여부 검증 시작]을 실행하세요</td></tr>
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
  const thead = document.getElementById('tableHead');
  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="11" class="empty-msg">데이터 없음</td></tr>';
    return;
  }
  const isAI = rows[0] && rows[0].mode === 'ai';
  if (isAI) {
    thead.innerHTML = '<tr><th>No</th><th>사용용도</th><th>품목</th><th>전표구분</th><th>승인번호</th><th>공급사업자</th><th>공급가액</th><th>매칭PDF</th><th>홈택스</th><th>증빙적합</th><th>AI검증</th><th>비고</th></tr>';
    tbody.innerHTML = rows.map(r => {
      let hcls = '';
      if (r.hometax === 'O') hcls = 'verdict-o';
      else if (r.hometax === 'X') hcls = 'verdict-x';
      let ecls = '';
      if (r.evidence_ok === 'O') ecls = 'verdict-o';
      else if (r.evidence_ok === 'X') ecls = 'verdict-x';
      let acls = '';
      if (r.ai_result === 'PASS') acls = 'suit-pass';
      else if (r.ai_result === 'FAIL') acls = 'suit-fail';
      else if (r.ai_result === 'REVIEW') acls = 'suit-review';
      let rcls = '';
      if (r.remark && r.remark.includes('통과')) rcls = 'remark-ok';
      else if (r.remark && r.remark.includes('누락')) rcls = 'remark-warn';
      let mcls = r.matched_pdf ? '' : 'verdict-q';
      return `<tr>
        <td>${r.no}</td><td>${esc(r.usage)}</td><td>${esc(r.item)}</td><td>${esc(r.doc_type)}</td>
        <td>${esc(r.approval)}</td><td>${esc(r.supplier_name)}</td><td class="amt">${r.amount}</td>
        <td class="${mcls}">${esc(r.matched_pdf) || '증빙없음'}</td><td class="${hcls}">${r.hometax || '-'}</td>
        <td class="${ecls}">${r.evidence_ok || '-'}</td>
        <td class="${acls}">${r.ai_result || '-'}</td><td class="${rcls}">${esc(r.remark) || '-'}</td>
      </tr>`;
    }).join('');
  } else {
    thead.innerHTML = '<tr><th>No</th><th>파일명</th><th>페이지</th><th>승인번호</th><th>공급자번호</th><th>작성일자</th><th>합계금액</th><th>진위여부</th><th>사용용도</th><th>적합성</th><th>비고</th></tr>';
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
      const suit = r.suitability || '';
      let scls = '';
      if (suit === 'PASS') scls = 'suit-pass';
      else if (suit === 'FAIL') scls = 'suit-fail';
      else if (suit === 'REVIEW') scls = 'suit-review';
      return `<tr class="${cls}">
        <td>${r.no}</td><td>${esc(r.filename)}</td><td>${r.page}</td>
        <td>${esc(r.approval)}</td><td>${r.supplier}</td><td>${r.date}</td>
        <td class="amt">${r.amount}</td><td class="${vcls}">${r.verdict || '-'}</td>
        <td>${esc(r.usage || '')}</td><td class="${scls}">${suit || '-'}</td><td class="${rcls}">${esc(remark) || '-'}</td>
      </tr>`;
    }).join('');
  }
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
    if (s.mode === 'ai') {
      area.innerHTML = `
        <span class="stat-chip">전체 ${s.total}건</span>
        <span class="stat-chip">매칭 ${s.matched}건</span>
        <span class="stat-chip ok">홈택스 O ${s.hometax_ok}건</span>
        <span class="stat-chip fail">홈택스 X ${s.hometax_x}건</span>
        <span class="stat-chip ok">증빙적합 O ${s.ev_ok}건</span>
        <span class="stat-chip fail">증빙적합 X ${s.ev_x}건</span>
        <span class="stat-chip ok">AI PASS ${s.ai_pass}건</span>
        <span class="stat-chip fail">AI FAIL ${s.ai_fail}건</span>
        <span class="stat-chip warn">AI REVIEW ${s.ai_review}건</span>
        <span class="stat-chip warn">증빙없음 ${s.no_evidence}건</span>
      `;
    } else {
      area.innerHTML = `
        <span class="stat-chip">전체 ${s.total}건</span>
        <span class="stat-chip">QR인식 ${s.qr_ok}건</span>
        <span class="stat-chip ok">O 정상 ${s.ok}건</span>
        <span class="stat-chip fail">X 불일치 ${s.fail}건</span>
        <span class="stat-chip warn">? 확인불가 ${s.unknown}건</span>
      `;
    }
  });
}

// ── 폴더 선택 ──────────────────────────────
function browseFolder() {
  fetch('/api/browse', {method:'POST'})
    .then(r => r.json())
    .then(d => { if (d.folder) document.getElementById('folderInput').value = d.folder; });
}

// ── 모드 & API Key ──────────────────────────
function getMode() {
  return document.querySelector('input[name="mode"]:checked').value;
}

function onModeChange() {
  const isAI = getMode() === 'ai';
  document.getElementById('apikeyRow').classList.toggle('show', isAI);
  document.getElementById('listUploadRow').classList.toggle('show', isAI);
  // AI 모드에서는 1단계/2단계 개별 버튼 숨김
  document.getElementById('btnExtract').style.display = isAI ? 'none' : '';
  document.getElementById('btnVerify').style.display = isAI ? 'none' : '';
  document.getElementById('btnResume').style.display = isAI ? 'none' : '';
  // 버튼 텍스트 변경
  document.getElementById('btnFullText').textContent = isAI ? 'AI 검증 시작' : '진위여부 검증 시작';
}

function setApiKey() {
  const key = document.getElementById('apikeyInput').value.trim();
  if (!key) { alert('API Key를 입력하세요.'); return; }
  fetch('/api/set-apikey', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({key})
  }).then(r => r.json()).then(d => {
    if (d.error) { alert(d.error); return; }
    const st = document.getElementById('apikeyStatus');
    st.textContent = '설정완료';
    st.className = 'apikey-status set';
  });
}

function uploadList() {
  const fileInput = document.getElementById('listFileInput');
  if (!fileInput.files.length) { alert('집행실적 엑셀 파일을 선택하세요.'); return; }
  const formData = new FormData();
  formData.append('file', fileInput.files[0]);
  fetch('/api/upload-list', {
    method: 'POST',
    body: formData
  }).then(r => r.json()).then(d => {
    if (d.error) { alert(d.error); return; }
    const st = document.getElementById('listStatus');
    st.textContent = d.count + '건 로드';
    st.className = 'list-status loaded';
  });
}

// ── 버튼 핸들러 ───────────────────────────────
function startFull() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) { alert('PDF 폴더 경로를 입력하세요.'); return; }
  const mode = getMode();
  const endpoint = mode === 'ai' ? '/api/full-ai' : '/api/full';
  fetch(endpoint, {
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
      document.getElementById('tableBody').innerHTML = '<tr><td colspan="11" class="empty-msg">PDF 폴더를 지정하고 [진위여부 검증 시작]을 실행하세요</td></tr>';
    })
    .finally(() => {
      btn.disabled = false;
      txt.textContent = '새로고침';
    });
}

function downloadExcel() {
  window.location.href = '/api/download';
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
