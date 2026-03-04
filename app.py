"""
전자세금계산서 QR 추출 + 홈택스 진위 확인 통합 웹 앱
=====================================================
건설현장 산업안전보건관리비 정산용
Streamlit 기반 단일 파일 웹 애플리케이션

실행 방법:
    streamlit run app.py

필수 라이브러리:
    pip install streamlit PyMuPDF pyzbar opencv-python-headless pandas openpyxl
    pip install pillow zxing-cpp selenium webdriver-manager requests
"""

import io
import re
import time
import random
import logging
import tempfile
import os
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote, urlencode

import streamlit as st
import pandas as pd
import numpy as np
import fitz                          # PyMuPDF
import cv2
from PIL import Image
from pyzbar import pyzbar
import zxingcpp

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
DPI = 300
ZOOM = DPI / 72

SLEEP_MIN = 3
SLEEP_MAX = 7
WAIT_TIMEOUT = 20

HOMETAX_MAIN = "https://www.hometax.go.kr"
MOBILE_API = "https://mob.tbet.hometax.go.kr/jsonAction.do"

ETAX_MENU_CANDIDATES = [
    "BC0201020200",
    "BC0101030000",
    "AB021",
    "AB0201",
]

SUCCESS_TEXTS = [
    "발급된 사실이 있습니다",
    "발급사실이 있습니다",
    "정상 발급",
    "발급 확인",
    "true",
    "0000",
]
FAILURE_TEXTS = [
    "발급 사실이 없습니다",
    "발급사실이 없습니다",
    "일치하지 않습니다",
    "조회되지 않습니다",
    "해당 자료가 없습니다",
    "없습니다",
    "false",
    "9999",
]
CAPTCHA_TEXTS = [
    "자동입력방지", "보안문자", "CAPTCHA",
    "captcha", "자동화", "확인코드",
]


# ══════════════════════════════════════════════
# 1. QR 데이터 파싱
# ══════════════════════════════════════════════
def parse_qr_data(raw: str) -> dict:
    """전자세금계산서 QR 문자열 → 딕셔너리 파싱 (URL / 파이프 형식)"""
    # ── 형식 1: 홈택스 URL ──
    if "hometax" in raw.lower() or raw.startswith("http"):
        try:
            decoded = unquote(unquote(raw))
            inner = decoded
            if "link=" in decoded:
                inner = decoded.split("link=", 1)[1]
            parsed = urlparse(inner)
            params = parse_qs(parsed.query)
            if "action" in params:
                action_url = unquote(params["action"][0])
                parsed2 = urlparse(action_url)
                params.update(parse_qs(parsed2.query))

            def _get(key):
                return params.get(key, [""])[0].strip()

            승인번호 = _get("etan")
            공급자번호 = _get("splrTxprDscmNo")
            작성일자 = _get("wrtDt")
            합계금액 = _get("splrCft")
            if 작성일자 and len(작성일자) == 8 and 작성일자.isdigit():
                작성일자 = f"{작성일자[:4]}-{작성일자[4:6]}-{작성일자[6:]}"
            if 승인번호:
                return {
                    "승인번호": 승인번호,
                    "공급자번호": 공급자번호,
                    "작성일자": 작성일자,
                    "합계금액": 합계금액,
                    "원본QR": raw,
                }
        except Exception:
            pass

    # ── 형식 2: 파이프(|) 구분자 ──
    parts = raw.split("|")
    if len(parts) >= 4:
        return {
            "승인번호": parts[0].strip(),
            "공급자번호": parts[1].strip(),
            "작성일자": parts[2].strip(),
            "합계금액": parts[3].strip(),
            "원본QR": raw,
        }

    # ── 형식 불명 ──
    return {
        "승인번호": raw.strip(),
        "공급자번호": "",
        "작성일자": "",
        "합계금액": "",
        "원본QR": raw,
    }


# ══════════════════════════════════════════════
# 2. 이미지에서 QR 코드 탐지 (다단계 엔진)
# ══════════════════════════════════════════════
def _pyzbar_decode(arr: np.ndarray) -> list[str]:
    results = pyzbar.decode(arr)
    return [
        obj.data.decode("utf-8", errors="replace")
        for obj in results if obj.type == "QRCODE"
    ]


def _zxing_decode(pil_img: Image.Image) -> list[str]:
    try:
        results = zxingcpp.read_barcodes(pil_img)
        return [r.text for r in results if "QR" in r.format.name]
    except Exception:
        return []


def detect_qr_from_image(img_array: np.ndarray) -> list[str]:
    """단일 이미지에서 QR 코드 추출 (pyzbar → ZXing, 전처리 반복)"""
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    else:
        gray = img_array.copy()

    pil_rgb = Image.fromarray(
        cv2.cvtColor(
            img_array if len(img_array.shape) == 3
            else cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR),
            cv2.COLOR_BGR2RGB,
        )
    )

    variants: list[tuple[str, np.ndarray | Image.Image, str]] = []
    variants.append(("원본", img_array, "pyzbar"))
    variants.append(("원본", pil_rgb, "zxing"))
    variants.append(("그레이", gray, "pyzbar"))

    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    sharpened = cv2.filter2D(gray, -1, kernel)
    variants.append(("샤프닝", sharpened, "pyzbar"))

    adaptive = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
    )
    variants.append(("적응형이진화", adaptive, "pyzbar"))
    variants.append(("적응형이진화", Image.fromarray(adaptive).convert("RGB"), "zxing"))

    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variants.append(("Otsu", otsu, "pyzbar"))

    enlarged = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    variants.append(("2배확대", enlarged, "pyzbar"))
    variants.append(("2배확대", Image.fromarray(enlarged).convert("RGB"), "zxing"))

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
# 3. 단일 PDF 처리 (BytesIO / 파일 경로 모두 지원)
# ══════════════════════════════════════════════
def process_pdf_bytes(pdf_bytes: bytes, filename: str) -> dict:
    """
    PDF 바이트에서 QR 코드를 탐지하고 첫 번째 결과를 반환.
    전체 페이지 → 3x3 분할 순으로 시도.
    """
    log.info(f"처리 중: {filename}")
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        log.error(f"  PDF 열기 실패: {e}")
        return _fail_record(filename, f"PDF 열기 오류: {e}")

    mat = fitz.Matrix(ZOOM, ZOOM)
    total_pages = len(doc)

    for page_no in range(total_pages):
        page = doc[page_no]
        log.info(f"  페이지 {page_no + 1}/{total_pages} 스캔")

        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        img_bytes = pix.tobytes("png")
        pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        arr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)

        # 전체 이미지 탐지
        qr_texts = detect_qr_from_image(arr)
        if qr_texts:
            doc.close()
            return _build_record_from_qr(qr_texts[0], filename, page_no + 1)

        # 3×3 분할 탐지
        h, w = arr.shape[:2]
        for row in range(3):
            for col in range(3):
                y1, y2 = row * h // 3, (row + 1) * h // 3
                x1, x2 = col * w // 3, (col + 1) * w // 3
                crop = arr[y1:y2, x1:x2]
                qr_texts = detect_qr_from_image(crop)
                if qr_texts:
                    log.info(f"  QR 발견 (페이지 {page_no+1}, 구역 {row},{col})")
                    doc.close()
                    return _build_record_from_qr(qr_texts[0], filename, page_no + 1)

    doc.close()
    log.warning(f"  QR 인식 불가: {filename}")
    return _fail_record(filename)


def _build_record_from_qr(raw: str, filename: str, page_no: int) -> dict:
    log.info(f"  QR 인식 성공 (페이지 {page_no}): {raw[:60]}...")
    parsed = parse_qr_data(raw)
    parsed["파일명"] = filename
    parsed["인식페이지"] = page_no
    return parsed


def _fail_record(filename: str, reason: str = "인식 불가") -> dict:
    return {
        "파일명": filename,
        "인식페이지": "-",
        "승인번호": reason,
        "공급자번호": "",
        "작성일자": "",
        "합계금액": "",
        "원본QR": "",
    }


# ══════════════════════════════════════════════
# 4. 데이터 전처리 유틸리티 (홈택스 검증용)
# ══════════════════════════════════════════════
def clean_approval_number(raw: str) -> str:
    s = str(raw).strip()
    s = re.sub(r'[a-zA-Z]+$', '', s)
    digits = re.sub(r'\D', '', s)
    return digits[:24] if len(digits) >= 24 else digits


def clean_date(raw: str) -> str:
    return re.sub(r'\D', '', str(raw))[:8]


def clean_number(raw) -> str:
    return re.sub(r'\D', '', str(raw))


# ══════════════════════════════════════════════
# 5. 브라우저 설정 (Headless)
# ══════════════════════════════════════════════
_DRIVER_CACHE = os.path.expanduser("~/.wdm/drivers/chromedriver/win64")


def find_chromedriver() -> str | None:
    import glob as _glob
    patterns = [
        os.path.join(_DRIVER_CACHE, "**", "chromedriver.exe"),
        os.path.join(_DRIVER_CACHE, "**", "chromedriver-win64", "chromedriver.exe"),
    ]
    for pat in patterns:
        found = _glob.glob(pat, recursive=True)
        if found:
            found.sort(reverse=True)
            return found[0]
    return None


def setup_driver():
    """Chrome WebDriver 초기화 (헤드리스 모드)"""
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service

    options = webdriver.ChromeOptions()
    # ★ 헤드리스 모드 (웹 서버 환경 필수)
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--lang=ko-KR")
    options.add_argument("--window-size=1280,900")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    )

    driver_path = find_chromedriver()
    if not driver_path:
        try:
            os.environ["WDM_SSL_VERIFY"] = "0"
            from webdriver_manager.chrome import ChromeDriverManager
            driver_path = ChromeDriverManager().install()
        except Exception as e:
            log.warning(f"webdriver-manager 실패: {e}")

    if driver_path:
        service = Service(driver_path)
    else:
        log.warning("ChromeDriver 경로 미발견 → PATH 내 chromedriver 사용 시도")
        service = Service()

    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"},
    )
    return driver


# ══════════════════════════════════════════════
# 6. 결과 판정 유틸리티
# ══════════════════════════════════════════════
def _judge_text(text: str) -> str:
    t = str(text)
    for p in CAPTCHA_TEXTS:
        if p.lower() in t.lower():
            log.warning("CAPTCHA 감지 — 헤드리스 모드에서는 자동 처리 불가")
            return "CAPTCHA"
    for p in SUCCESS_TEXTS:
        if p in t:
            return "O"
    for p in FAILURE_TEXTS:
        if p in t:
            return "X"
    return ""


def _judge_page_text(driver) -> str:
    from selenium.webdriver.common.by import By

    texts = []
    try:
        texts.append(driver.find_element(By.TAG_NAME, "body").text)
    except Exception:
        pass
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.frame(iframe)
            texts.append(driver.find_element(By.TAG_NAME, "body").text)
            driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()
    return _judge_text("\n".join(texts))


# ══════════════════════════════════════════════
# 7. 전략 1: QR URL 직접 접근 (Selenium)
# ══════════════════════════════════════════════
def _extract_inner_url(qr_url: str) -> str:
    try:
        decoded = unquote(unquote(qr_url))
        if "link=" in decoded:
            inner = decoded.split("link=", 1)[1]
            if "action=" in inner:
                action_part = inner.split("action=", 1)[1]
                return unquote(action_part.split("&")[0])
            return inner
        if "hometax.go.kr" in qr_url or "mob.tbet" in qr_url:
            return qr_url
    except Exception:
        pass
    return ""


def verify_via_qr_url(driver, qr_url: str, row_data: dict) -> str:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    inner_url = _extract_inner_url(qr_url)
    if not inner_url:
        log.info("  [전략1] 내부 URL 추출 실패")
        return ""

    log.info(f"  [전략1] 내부 URL 접근: {inner_url[:80]}...")
    try:
        driver.get(inner_url)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(4)
        result = _judge_page_text(driver)
        if result:
            log.info(f"  [전략1] 판정: {result}")
        return result
    except Exception as e:
        log.warning(f"  [전략1] 오류: {e}")
        return ""


# ══════════════════════════════════════════════
# 8. 전략 2: 모바일 API 직접 호출 (requests)
# ══════════════════════════════════════════════
def verify_via_api(row_data: dict) -> str:
    import requests

    params = {
        "actionId": "UTBETBDA16F001",
        "menuId": "6001020100",
        "etan": row_data["승인번호_원본"],
        "splrTxprDscmNo": row_data["공급자번호"],
        "wrtDt": row_data["작성일자_raw"],
        "splrCft": row_data["합계금액"],
    }
    if row_data.get("수급자번호"):
        params["dmnrTxprDscmNo"] = row_data["수급자번호"]

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Linux; Android 13; SM-G991B) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/131.0.0.0 Mobile Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://mob.tbet.hometax.go.kr/",
        "Origin": "https://mob.tbet.hometax.go.kr",
    }

    log.info(f"  [전략2] 모바일 API 호출")
    try:
        resp = requests.get(
            MOBILE_API, params=params, headers=headers,
            timeout=15, verify=False,
        )
        log.info(f"  [전략2] 응답 코드: {resp.status_code}")

        text = resp.text
        try:
            data = resp.json()
            result_code = data.get("resultCode") or data.get("errCd") or data.get("result") or ""
            result_msg = data.get("resultMsg") or data.get("errMsg") or data.get("message") or ""
            text = str(result_code) + " " + str(result_msg)
        except Exception:
            pass

        return _judge_text(text)
    except Exception as e:
        log.warning(f"  [전략2] API 호출 오류: {e}")
        return ""


# ══════════════════════════════════════════════
# 9. 전략 3: 홈택스 WebSquare 메뉴 JS 내비게이션
# ══════════════════════════════════════════════
def verify_via_websquare(driver, row_data: dict) -> str:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    log.info("  [전략3] 홈택스 WebSquare 내비게이션 시작")

    driver.get(HOMETAX_MAIN)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(5)

    nav_success = False
    for menu_cd in ETAX_MENU_CANDIDATES:
        try:
            driver.execute_script(f"gfn_viewMenu('{menu_cd}')")
            time.sleep(3)
            body_text = driver.find_element(By.TAG_NAME, "body").text
            if "전자세금계산서" in body_text or "승인번호" in body_text:
                log.info(f"  [전략3] 메뉴 이동 성공 (menuCd={menu_cd})")
                nav_success = True
                break
        except Exception:
            pass

    if not nav_success:
        for menu_cd in ETAX_MENU_CANDIDATES:
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    try:
                        driver.switch_to.frame(iframe)
                        driver.execute_script(f"gfn_viewMenu('{menu_cd}')")
                        time.sleep(2)
                        body_text = driver.find_element(By.TAG_NAME, "body").text
                        if "전자세금계산서" in body_text or "승인번호" in body_text:
                            nav_success = True
                            break
                        driver.switch_to.default_content()
                    except Exception:
                        driver.switch_to.default_content()
                if nav_success:
                    break
            except Exception:
                pass

    if not nav_success:
        log.warning("  [전략3] 자동 메뉴 이동 실패")
        return "확인불가"

    return _fill_websquare_form(driver, row_data)


def _fill_websquare_form(driver, row_data: dict) -> str:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import NoSuchElementException

    field_map = {
        "승인번호": (row_data["승인번호_정제"], ["etan", "tbx_etan", "aprvNo", "inp_etan"]),
        "공급자번호": (row_data["공급자번호"], ["splrTxprDscmNo", "tbx_splrBizNo", "inp_splrBizNo"]),
        "작성일자": (row_data["작성일자_raw"], ["wrtDt", "tbx_wrtDt", "inp_wrtDt"]),
        "합계금액": (row_data["합계금액"], ["splrCft", "tbx_splrCft", "inp_totAmt"]),
    }

    def _try_fill(candidates, value):
        for ctx in [None] + driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                if ctx:
                    driver.switch_to.frame(ctx)
                for cid in candidates:
                    for method in [By.ID, By.NAME]:
                        try:
                            el = driver.find_element(method, cid)
                            if el.is_displayed():
                                el.click()
                                el.send_keys(Keys.CONTROL + "a")
                                el.send_keys(value)
                                if ctx:
                                    driver.switch_to.default_content()
                                return True
                        except NoSuchElementException:
                            pass
                if ctx:
                    driver.switch_to.default_content()
            except Exception:
                driver.switch_to.default_content()
        return False

    success = 0
    for label, (value, candidates) in field_map.items():
        if _try_fill(candidates, value):
            log.info(f"    입력 [{label}]: {value}")
            success += 1
            time.sleep(0.4)
        else:
            log.warning(f"    입력 필드 미발견: {label}")

    driver.switch_to.default_content()

    if success == 0:
        return "확인불가"

    _click_search_button(driver)
    time.sleep(3)

    return _judge_page_text(driver) or "확인불가"


def _click_search_button(driver):
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException

    btn_xpaths = [
        '//button[contains(text(),"조회하기")]',
        '//button[contains(text(),"조회")]',
        '//input[@type="button" and contains(@value,"조회")]',
        '//a[contains(text(),"조회")]',
    ]
    for ctx in [None] + driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            if ctx:
                driver.switch_to.frame(ctx)
            for xp in btn_xpaths:
                try:
                    btn = driver.find_element(By.XPATH, xp)
                    if btn.is_displayed():
                        driver.execute_script("arguments[0].click();", btn)
                        log.info("    조회 버튼 클릭 성공")
                        if ctx:
                            driver.switch_to.default_content()
                        return
                except NoSuchElementException:
                    pass
            if ctx:
                driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()
    log.warning("    조회 버튼 미발견")


# ══════════════════════════════════════════════
# 10. 단건 조회 (3중 전략 통합) — dict 입력 버전
# ══════════════════════════════════════════════
def verify_one(driver, record: dict) -> str:
    """
    단건 전자세금계산서 조회 (3중 전략).
    record: process_pdf_bytes()가 반환한 딕셔너리.
    """
    row_data = {
        "승인번호_원본": str(record.get("승인번호", "")),
        "승인번호_정제": clean_approval_number(str(record.get("승인번호", ""))),
        "공급자번호": clean_number(record.get("공급자번호", "")),
        "작성일자_raw": clean_date(str(record.get("작성일자", ""))),
        "합계금액": clean_number(record.get("합계금액", "")),
        "수급자번호": clean_number(record.get("수급자번호", "")),
    }
    qr_url = str(record.get("원본QR", ""))

    log.info(
        f"  조회 데이터: 승인번호={row_data['승인번호_정제']} "
        f"공급자={row_data['공급자번호']} "
        f"일자={row_data['작성일자_raw']} "
        f"금액={row_data['합계금액']}"
    )

    # 전략 1: QR URL 직접 접근
    if qr_url and qr_url.startswith("http"):
        result = verify_via_qr_url(driver, qr_url, row_data)
        if result and result != "CAPTCHA":
            return result

    # 전략 2: 모바일 API 직접 호출
    result = verify_via_api(row_data)
    if result and result != "CAPTCHA":
        return result

    # 전략 3: WebSquare 메뉴 내비게이션
    result = verify_via_websquare(driver, row_data)
    return result if result else "확인불가"


# ══════════════════════════════════════════════
# 11. 엑셀 다운로드 바이트 생성
# ══════════════════════════════════════════════
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame → Excel 바이트 (다운로드용)"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="진위확인결과")
        ws = writer.sheets["진위확인결과"]
        for col in ws.columns:
            max_len = max(
                (len(str(c.value)) for c in col if c.value is not None),
                default=8,
            )
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    return buf.getvalue()


# ══════════════════════════════════════════════
# 12. Streamlit UI
# ══════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="전자세금계산서 진위 확인",
        page_icon="📄",
        layout="wide",
    )

    # ══════════════════════════════════════════════
    # 사이드바 안내 패널
    # ══════════════════════════════════════════════
    with st.sidebar:
        st.markdown("## 🏗️ 앱 소개")
        st.info(
            "**건설현장 산업안전보건관리비 정산**을 위한\n"
            "전자세금계산서 자동 검증 도구입니다.\n\n"
            "PDF를 업로드하면 QR 코드를 자동 추출하고,\n"
            "홈택스를 통해 진위 여부를 확인합니다."
        )

        st.markdown("---")
        st.markdown("## ⚙️ 동작 매커니즘")

        # ── 1단계 ──
        with st.expander("🔍 1단계 — QR 정밀 추출", expanded=True):
            st.markdown(
                "업로드된 PDF를 **고해상도(300 DPI) 이미지**로 변환한 뒤,\n"
                "**2개의 인식 엔진**을 동시 가동하여 QR을 탐색합니다.\n\n"
                "| 엔진 | 특징 |\n"
                "|------|------|\n"
                "| **pyzbar** | 빠른 1차 스캔 |\n"
                "| **ZXing-C++** | 손상 QR 보정 인식 |\n\n"
                "한 번에 인식되지 않으면 **6가지 이미지 전처리**를\n"
                "자동으로 적용하여 재시도합니다.\n\n"
                "- 🖼️ 그레이스케일 변환\n"
                "- 🔪 샤프닝 필터\n"
                "- 🎛️ 적응형 이진화 (Adaptive Threshold)\n"
                "- 📊 Otsu 이진화\n"
                "- 🔎 2배 확대 보간\n"
                "- 🧩 **3×3 분할 스캔** — 페이지를 9개 구역으로\n"
                "  나눠 구석에 있는 QR도 끈질기게 찾아냅니다."
            )

        # ── 2단계 ──
        with st.expander("🛡️ 2단계 — 홈택스 3중 검증"):
            st.markdown(
                "추출된 **승인번호 · 공급자번호 · 작성일자 · 합계금액**을\n"
                "바탕으로, 3가지 전략을 순차적으로 시도합니다.\n"
            )
            st.markdown(
                "**① QR 내부 URL 직접 접근**\n"
                "> QR에 포함된 홈택스 검증 URL로 바로 접속하여\n"
                "> 발급 사실을 즉시 확인합니다.\n\n"
                "**② 모바일 API 호출**\n"
                "> 홈택스 모바일 JSON API를 직접 호출하여\n"
                "> 빠르고 안정적으로 결과를 조회합니다.\n\n"
                "**③ WebSquare 브라우저 자동화**\n"
                "> Selenium 헤드리스 브라우저가 홈택스에\n"
                "> 직접 접속 → 메뉴 탐색 → 값 입력 → 조회를\n"
                "> 자동 수행합니다."
            )
            st.warning(
                "⏱️ 홈택스 과부하 방지를 위해 건별로\n"
                "3~7초의 랜덤 대기 시간이 적용됩니다."
            )

        # ── 3단계 ──
        with st.expander("📊 3단계 — 결과 도출"):
            st.markdown(
                "각 세금계산서에 대해 아래와 같이 판정합니다.\n\n"
                "| 판정 | 의미 |\n"
                "|------|------|\n"
                "| ✅ **O** | 정상 — 발급 사실 확인 |\n"
                "| ❌ **X** | 불일치 — 발급 사실 없음 |\n"
                "| ⚠️ **확인불가** | 홈택스 접속 실패 등 |\n"
                "| 🔒 **CAPTCHA** | 보안문자 감지 |\n\n"
                "검증 완료 후 **엑셀 파일(.xlsx)**로\n"
                "결과를 다운로드할 수 있습니다."
            )

        st.markdown("---")
        st.markdown(
            "<div style='text-align:center; color:gray; font-size:0.85em;'>"
            "📄 전자세금계산서 진위 확인 v1.0<br>"
            "Streamlit 기반 웹 애플리케이션"
            "</div>",
            unsafe_allow_html=True,
        )

    # ══════════════════════════════════════════════
    # 메인 영역
    # ══════════════════════════════════════════════
    st.title("📄 전자세금계산서 QR 추출 & 홈택스 진위 확인")
    st.caption("건설현장 산업안전보건관리비 정산용 — PDF 업로드 → QR 추출 → 홈택스 자동 검증")

    # ── SSL 경고 억제 ──
    import urllib3
    urllib3.disable_warnings()

    # ── 파일 업로드 ──
    uploaded_files = st.file_uploader(
        "전자세금계산서 PDF 파일을 업로드하세요 (여러 파일 선택 가능)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("PDF 파일을 업로드하면 QR 코드 추출 및 홈택스 진위 확인을 진행합니다.")
        return

    st.success(f"{len(uploaded_files)}개 파일이 업로드되었습니다.")

    # ── 검증 시작 버튼 ──
    if not st.button("🔍 진위 검증 시작", type="primary", use_container_width=True):
        return

    # ── 결과 테이블 영역 ──
    col_order = ["파일명", "인식페이지", "승인번호", "공급자번호", "작성일자", "합계금액", "진위여부"]
    results_placeholder = st.empty()
    status_text = st.empty()
    progress_bar = st.progress(0)

    records: list[dict] = []
    total = len(uploaded_files)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Phase 1: QR 추출
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.subheader("Phase 1 — QR 코드 추출")

    for i, uploaded in enumerate(uploaded_files):
        fname = uploaded.name
        status_text.markdown(f"**[{i+1}/{total}]** `{fname}` — QR 추출 중...")

        pdf_bytes = uploaded.read()
        record = process_pdf_bytes(pdf_bytes, fname)
        record["진위여부"] = ""  # 아직 미검증
        records.append(record)

        # 실시간 테이블 갱신
        df_display = pd.DataFrame(records, columns=col_order)
        results_placeholder.dataframe(df_display, use_container_width=True, hide_index=True)
        progress_bar.progress((i + 1) / total * 0.5)  # 0~50%

    status_text.markdown("**QR 추출 완료!** 홈택스 진위 확인을 시작합니다...")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Phase 2: 홈택스 진위 확인
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.subheader("Phase 2 — 홈택스 진위 확인")

    # 검증 대상 필터 (QR 인식 성공 건만)
    verify_indices = [
        j for j, r in enumerate(records)
        if r["승인번호"] not in ("인식 불가", "") and not r["승인번호"].startswith("PDF 열기 오류")
    ]

    if not verify_indices:
        status_text.warning("QR 인식에 성공한 건이 없어 홈택스 검증을 건너뜁니다.")
        progress_bar.progress(1.0)
    else:
        # Selenium 드라이버 초기화
        driver = None
        try:
            with st.spinner("Chrome 헤드리스 브라우저 시작 중..."):
                driver = setup_driver()

            for step, j in enumerate(verify_indices):
                rec = records[j]
                fname = rec["파일명"]
                status_text.markdown(
                    f"**[{step+1}/{len(verify_indices)}]** `{fname}` — 홈택스 검증 중..."
                )

                result = verify_one(driver, rec)
                rec["진위여부"] = result

                # 실시간 테이블 갱신
                df_display = pd.DataFrame(records, columns=col_order)
                results_placeholder.dataframe(df_display, use_container_width=True, hide_index=True)
                progress_bar.progress(0.5 + (step + 1) / len(verify_indices) * 0.5)

                # 홈택스 과부하 방지 대기
                if step < len(verify_indices) - 1:
                    wait_sec = random.uniform(SLEEP_MIN, SLEEP_MAX)
                    status_text.markdown(
                        f"**[{step+1}/{len(verify_indices)}]** `{fname}` — "
                        f"판정: **{result}** | 다음 건까지 {wait_sec:.1f}초 대기..."
                    )
                    time.sleep(wait_sec)
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    # QR 인식 실패 건은 진위여부를 '-'로 표시
    for rec in records:
        if not rec["진위여부"]:
            rec["진위여부"] = "-"

    progress_bar.progress(1.0)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 최종 결과 표시
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.subheader("최종 결과")
    df_final = pd.DataFrame(records, columns=col_order)
    results_placeholder.dataframe(df_final, use_container_width=True, hide_index=True)

    # 요약 통계
    ok = (df_final["진위여부"] == "O").sum()
    ng = (df_final["진위여부"] == "X").sum()
    unk = df_final["진위여부"].isin(["확인불가", "CAPTCHA"]).sum()
    skip = df_final["진위여부"].isin(["-", ""]).sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("정상 (O)", f"{ok}건")
    c2.metric("불일치 (X)", f"{ng}건")
    c3.metric("확인불가", f"{unk}건")
    c4.metric("미검증 / 인식실패", f"{skip}건")

    status_text.success(
        f"전체 {total}건 처리 완료 — O: {ok} / X: {ng} / 확인불가: {unk} / 미검증: {skip}"
    )

    # ── 엑셀 다운로드 ──
    excel_bytes = to_excel_bytes(df_final)
    st.download_button(
        label="📥 결과 엑셀 다운로드 (verification_result.xlsx)",
        data=excel_bytes,
        file_name="verification_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
