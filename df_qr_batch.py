"""
전자세금계산서 QR 코드 일괄 추출기
=====================================
건설현장 산업안전보건관리비 정산용
PDF 파일에서 QR 코드를 자동 인식하여 Excel로 저장

필수 라이브러리 설치:
    pip install PyMuPDF pyzbar opencv-python-headless pandas openpyxl pillow zxing-cpp

실행 방법:
    python df_qr_batch.py                     # 스크립트와 같은 폴더의 PDF 처리
    python df_qr_batch.py "C:/path/to/folder" # 지정 폴더의 PDF 처리
"""

import os
import sys
import io
import logging
from pathlib import Path

import fitz                        # PyMuPDF: PDF → 이미지 변환
import cv2                         # OpenCV: 이미지 전처리
import numpy as np
from PIL import Image              # Pillow: 이미지 객체 변환
from pyzbar import pyzbar          # pyzbar: QR 1차 디코딩
import zxingcpp                    # ZXing C++: QR 2차 디코딩 (더 강력)
import pandas as pd                # 데이터프레임 → Excel 저장

# ──────────────────────────────────────────────
# 로깅 설정
# ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────
DPI = 300            # PDF → 이미지 변환 해상도 (높을수록 인식률 ↑)
ZOOM = DPI / 72      # PyMuPDF 기본 72 DPI 기준 배율
OUTPUT_EXCEL = "results.xlsx"


# ══════════════════════════════════════════════
# 1. QR 데이터 파싱
# ══════════════════════════════════════════════
def parse_qr_data(raw: str) -> dict:
    """
    전자세금계산서 QR 문자열을 파싱하여 딕셔너리로 반환.

    지원 형식:
        1) 국세청 홈택스 URL 형식 (hometax.page.link 또는 hometax.go.kr 포함)
           URL 쿼리 파라미터에서 승인번호·공급자번호·작성일자·합계금액 추출.
           주요 파라미터:
             etan           → 승인번호 (전자세금계산서 승인번호)
             splrTxprDscmNo → 공급자번호
             wrtDt          → 작성일자
             splrCft        → 합계금액

        2) 파이프(|) 구분자 형식:
             승인번호|공급자번호|작성일자|합계금액|...

    Parameters
    ----------
    raw : str
        디코딩된 원시 QR 문자열

    Returns
    -------
    dict
        파싱된 4개 항목 + 원본 문자열
    """
    from urllib.parse import urlparse, parse_qs, unquote

    # ── 형식 1: 홈택스 URL ─────────────────────────
    if "hometax" in raw.lower() or raw.startswith("http"):
        try:
            # 중첩 인코딩 해제 (link=... 파라미터 내 이중 인코딩)
            decoded = unquote(unquote(raw))
            # 내부 URL 추출 (link= 이후 부분)
            inner = decoded
            if "link=" in decoded:
                inner = decoded.split("link=", 1)[1]

            parsed = urlparse(inner)
            params = parse_qs(parsed.query)

            # action 파라미터가 있으면 한 번 더 파싱
            if "action" in params:
                action_url = unquote(params["action"][0])
                parsed2 = urlparse(action_url)
                params.update(parse_qs(parsed2.query))

            def _get(key: str) -> str:
                return params.get(key, [""])[0].strip()

            승인번호   = _get("etan")
            공급자번호 = _get("splrTxprDscmNo")
            작성일자   = _get("wrtDt")
            합계금액   = _get("splrCft")

            # 작성일자 포맷 정리: 20250602 → 2025-06-02
            if 작성일자 and len(작성일자) == 8 and 작성일자.isdigit():
                작성일자 = f"{작성일자[:4]}-{작성일자[4:6]}-{작성일자[6:]}"

            if 승인번호:   # 필수 항목이 파싱된 경우만 사용
                return {
                    "승인번호":   승인번호,
                    "공급자번호": 공급자번호,
                    "작성일자":   작성일자,
                    "합계금액":   합계금액,
                    "원본QR":     raw,
                }
        except Exception:
            pass  # 파싱 실패 시 아래 파이프 방식으로 계속

    # ── 형식 2: 파이프(|) 구분자 ──────────────────
    parts = raw.split("|")
    if len(parts) >= 4:
        return {
            "승인번호":   parts[0].strip(),
            "공급자번호": parts[1].strip(),
            "작성일자":   parts[2].strip(),
            "합계금액":   parts[3].strip(),
            "원본QR":     raw,
        }

    # ── 형식 불명 (URL 전체를 승인번호 셀에 보존) ──
    return {
        "승인번호":   raw.strip(),
        "공급자번호": "",
        "작성일자":   "",
        "합계금액":   "",
        "원본QR":     raw,
    }


# ══════════════════════════════════════════════
# 2. 이미지에서 QR 코드 탐지 (3단계 엔진)
# ══════════════════════════════════════════════
def _pyzbar_decode(arr: np.ndarray) -> list[str]:
    """pyzbar로 디코딩 후 QR CODE 타입만 반환"""
    results = pyzbar.decode(arr)
    return [
        obj.data.decode("utf-8", errors="replace")
        for obj in results
        if obj.type == "QRCODE"
    ]


def _zxing_decode(pil_img: Image.Image) -> list[str]:
    """ZXing C++로 디코딩 (pyzbar가 실패한 경우 대비)"""
    try:
        results = zxingcpp.read_barcodes(pil_img)
        return [
            r.text for r in results
            if "QR" in r.format.name
        ]
    except Exception:
        return []


def detect_qr_from_image(img_array: np.ndarray) -> list[str]:
    """
    단일 이미지(numpy BGR array)에서 QR 코드를 추출.
    pyzbar → ZXing 순으로 시도하며, 전처리도 반복 적용.

    Parameters
    ----------
    img_array : np.ndarray
        BGR 또는 GRAY 형식의 이미지 배열

    Returns
    -------
    list[str]
        디코딩된 QR 문자열 목록 (없으면 빈 리스트)
    """
    # 그레이스케일 준비
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    else:
        gray = img_array.copy()

    # PIL 이미지 (ZXing용)
    pil_rgb = Image.fromarray(cv2.cvtColor(img_array if len(img_array.shape)==3 else
                                            cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR),
                                            cv2.COLOR_BGR2RGB))

    # ── 전처리 목록 ──────────────────────────────
    variants: list[tuple[str, np.ndarray | Image.Image, str]] = []

    # 1) 원본
    variants.append(("원본", img_array, "pyzbar"))
    variants.append(("원본", pil_rgb, "zxing"))

    # 2) 그레이스케일
    variants.append(("그레이", gray, "pyzbar"))

    # 3) 샤프닝
    kernel = np.array([[0,-1,0],[-1,5,-1],[0,-1,0]])
    sharpened = cv2.filter2D(gray, -1, kernel)
    variants.append(("샤프닝", sharpened, "pyzbar"))

    # 4) 적응형 이진화
    adaptive = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
    )
    variants.append(("적응형이진화", adaptive, "pyzbar"))
    variants.append(("적응형이진화", Image.fromarray(adaptive).convert("RGB"), "zxing"))

    # 5) Otsu 이진화
    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variants.append(("Otsu", otsu, "pyzbar"))

    # 6) 2배 확대 (소형 QR 대응)
    enlarged = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    variants.append(("2배확대", enlarged, "pyzbar"))
    variants.append(("2배확대", Image.fromarray(enlarged).convert("RGB"), "zxing"))

    # ── 탐지 실행 ────────────────────────────────
    for name, arr_or_pil, engine in variants:
        if engine == "pyzbar":
            found = _pyzbar_decode(arr_or_pil)
        else:
            found = _zxing_decode(arr_or_pil)
        if found:
            log.debug(f"    [{engine}/{name}] 인식 성공")
            return found

    return []


# ══════════════════════════════════════════════
# 3. 단일 PDF 처리
# ══════════════════════════════════════════════
def process_pdf(pdf_path: str) -> dict:
    """
    단일 PDF의 모든 페이지를 고해상도로 렌더링하여
    QR 코드를 탐지하고 첫 번째 결과를 반환.

    전략:
        1) 전체 페이지 이미지로 탐지
        2) 실패 시 3×3 분할 탐지 (QR이 구석에 있어도 인식)

    Parameters
    ----------
    pdf_path : str
        처리할 PDF 파일 경로

    Returns
    -------
    dict
        파싱된 QR 정보 또는 '인식 불가' 표시 딕셔너리
    """
    filename = Path(pdf_path).name
    log.info(f"처리 중: {filename}")

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        log.error(f"  PDF 열기 실패: {e}")
        return _fail_record(filename, f"PDF 열기 오류: {e}")

    mat = fitz.Matrix(ZOOM, ZOOM)
    total_pages = len(doc)

    for page_no in range(total_pages):
        page = doc[page_no]
        log.info(f"  페이지 {page_no + 1}/{total_pages} 스캔")

        # PDF 페이지 → RGB PIL → BGR numpy
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        img_bytes = pix.tobytes("png")
        pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        arr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)

        # ── 전체 이미지 탐지 ──
        qr_texts = detect_qr_from_image(arr)
        if qr_texts:
            return _build_record(qr_texts[0], filename, page_no + 1, doc)

        # ── 3×3 분할 탐지 ──
        h, w = arr.shape[:2]
        for row in range(3):
            for col in range(3):
                y1, y2 = row * h // 3, (row + 1) * h // 3
                x1, x2 = col * w // 3, (col + 1) * w // 3
                crop = arr[y1:y2, x1:x2]
                qr_texts = detect_qr_from_image(crop)
                if qr_texts:
                    log.info(f"  QR 발견 (페이지 {page_no+1}, 구역 {row},{col})")
                    return _build_record(qr_texts[0], filename, page_no + 1, doc)

    doc.close()
    log.warning(f"  QR 인식 불가: {filename}")
    return _fail_record(filename)


def _build_record(raw: str, filename: str, page_no: int, doc) -> dict:
    """QR 인식 성공 시 레코드 생성"""
    log.info(f"  QR 인식 성공 (페이지 {page_no}): {raw[:60]}...")
    doc.close()
    parsed = parse_qr_data(raw)
    parsed["파일명"] = filename
    parsed["인식페이지"] = page_no
    return parsed


def _fail_record(filename: str, reason: str = "인식 불가") -> dict:
    """QR 인식 실패 시 빈 레코드 반환"""
    return {
        "파일명":    filename,
        "인식페이지": "-",
        "승인번호":   reason,
        "공급자번호":  "",
        "작성일자":   "",
        "합계금액":   "",
        "원본QR":    "",
    }


# ══════════════════════════════════════════════
# 4. 일괄 처리 (Batch Process)
# ══════════════════════════════════════════════
def batch_process(folder_path: str) -> None:
    """
    지정 폴더 내 모든 PDF를 일괄 처리하고 results.xlsx 저장.

    Parameters
    ----------
    folder_path : str
        PDF 파일들이 담긴 폴더 경로
    """
    folder = Path(folder_path).resolve()
    if not folder.exists():
        log.error(f"폴더를 찾을 수 없습니다: {folder}")
        sys.exit(1)

    pdf_files = sorted(folder.glob("*.pdf"))
    if not pdf_files:
        log.warning(f"'{folder}' 안에 PDF 파일이 없습니다.")
        return

    log.info(f"총 {len(pdf_files)}개 PDF 파일 처리 시작")
    log.info("=" * 55)

    records: list[dict] = []
    for idx, pdf_path in enumerate(pdf_files, start=1):
        log.info(f"[{idx}/{len(pdf_files)}] {pdf_path.name}")
        records.append(process_pdf(str(pdf_path)))
        log.info("-" * 55)

    # ── Excel 저장 ──────────────────────────────
    col_order = ["파일명", "인식페이지", "승인번호", "공급자번호", "작성일자", "합계금액", "원본QR"]
    df = pd.DataFrame(records, columns=col_order)

    output_path = folder / OUTPUT_EXCEL
    df.to_excel(output_path, index=False, engine="openpyxl")
    _auto_fit_excel(output_path)

    success = (df["승인번호"].notna() & ~df["승인번호"].isin(["인식 불가"])).sum()
    log.info(f"저장 완료 → {output_path}")
    log.info(f"성공: {success} / 전체: {len(df)}")


def _auto_fit_excel(path: Path) -> None:
    """Excel 열 너비를 내용 길이에 맞게 자동 조정"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active
        for col in ws.columns:
            max_len = max(
                (len(str(cell.value)) for cell in col if cell.value is not None),
                default=8,
            )
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
        wb.save(path)
    except Exception as e:
        log.warning(f"열 너비 자동 조정 실패 (무시): {e}")


# ══════════════════════════════════════════════
# 5. 진입점
# ══════════════════════════════════════════════
if __name__ == "__main__":
    # 실행 방법 1: python df_qr_batch.py "C:/path/to/pdf_folder"
    # 실행 방법 2: python df_qr_batch.py  (스크립트와 같은 폴더 처리)
    target_folder = sys.argv[1] if len(sys.argv) >= 2 else str(Path(__file__).parent)
    batch_process(target_folder)
