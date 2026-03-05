"""
전자세금계산서 진위 확인 자동화 (홈택스)
==========================================
건설현장 산업안전보건관리비 정산용
results.xlsx의 승인번호를 홈택스에서 자동 조회하여
진위 여부(O/X/확인불가)를 verification_final.xlsx에 저장

필수 라이브러리 설치:
    pip install selenium webdriver-manager pandas openpyxl requests

실행 방법:
    python hometax_verify.py

동작 방식 (3중 전략):
    1단계 (기본) : QR 원본 URL → 홈택스 모바일 웹 결과 페이지 접근
    2단계 (대체1): 모바일 API 직접 호출 (requests)
    3단계 (대체2): 홈택스 WebSquare 메뉴 JS 내비게이션 + 폼 입력
"""

import os, re, sys, time, random, json, logging, ssl, urllib.request
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote, urlencode

import pandas as pd

# ──────────────────────────────────────────────
# 로깅
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
INPUT_FILE   = "results.xlsx"
OUTPUT_FILE  = "verification_final.xlsx"
SLEEP_MIN    = 3
SLEEP_MAX    = 7
WAIT_TIMEOUT = 20

# 기기에 설치된 ChromeDriver 경로 (버전 자동 감지)
_DRIVER_CACHE = os.path.expanduser(
    "~/.wdm/drivers/chromedriver/win64"
)

# 홈택스 관련 URL
HOMETAX_MAIN = "https://www.hometax.go.kr"
MOBILE_API   = "https://mob.tbet.hometax.go.kr/jsonAction.do"

# 비로그인 전자세금계산서 조회 menuCd 후보 (홈택스 버전별 상이)
ETAX_MENU_CANDIDATES = [
    "BC0201020200",  # 전자세금계산서 조회
    "BC0101030000",
    "AB021",
    "AB0201",
]


# ══════════════════════════════════════════════
# 1. 데이터 전처리 유틸리티
# ══════════════════════════════════════════════
def clean_approval_number(raw: str) -> str:
    """
    승인번호 정제.
    'cr', 'b' 등 QR URL 파라미터 접미사 제거 후 24자리 숫자 추출.
    숫자가 24자 미만이면 가능한 전체 반환.
    """
    s = str(raw).strip()
    s = re.sub(r'[a-zA-Z]+$', '', s)          # 끝 영문자 제거
    digits = re.sub(r'\D', '', s)              # 숫자만 추출
    return digits[:24] if len(digits) >= 24 else digits


def clean_date(raw: str) -> str:
    """작성일자 → YYYYMMDD (8자리 숫자)"""
    return re.sub(r'\D', '', str(raw))[:8]


def clean_number(raw) -> str:
    """금액·사업자번호 등 숫자만 추출"""
    return re.sub(r'\D', '', str(raw))


def find_chromedriver() -> str | None:
    """설치된 chromedriver.exe 경로 자동 탐색"""
    import glob
    patterns = [
        os.path.join(_DRIVER_CACHE, "**", "chromedriver.exe"),
        os.path.join(_DRIVER_CACHE, "**", "chromedriver-win64", "chromedriver.exe"),
    ]
    for pat in patterns:
        found = glob.glob(pat, recursive=True)
        if found:
            # 버전 내림차순 정렬 → 가장 높은 버전 선택
            found.sort(reverse=True)
            return found[0]
    return None


# ══════════════════════════════════════════════
# 2. 브라우저 설정
# ══════════════════════════════════════════════
def setup_driver():
    """
    Chrome WebDriver 초기화.
    webdriver-manager 네트워크 오류 시 로컬 캐시 또는 PATH 내 드라이버 사용.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service

    options = webdriver.ChromeOptions()
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

    # ── 드라이버 경로 탐색 순서 ──────────────────────
    # 1) 로컬 캐시에서 찾기
    driver_path = find_chromedriver()

    # 2) webdriver-manager (SSL 우회)
    if not driver_path:
        try:
            os.environ["WDM_SSL_VERIFY"] = "0"
            from webdriver_manager.chrome import ChromeDriverManager
            driver_path = ChromeDriverManager().install()
        except Exception as e:
            log.warning(f"webdriver-manager 실패: {e}")

    # 3) PATH에서 chromedriver 사용
    if driver_path:
        service = Service(driver_path)
    else:
        log.warning("ChromeDriver 경로 미발견 → PATH 내 chromedriver 사용 시도")
        service = Service()

    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"}
    )
    return driver


# ══════════════════════════════════════════════
# 3. 전략 1: QR URL 직접 접근 (Selenium)
# ══════════════════════════════════════════════
def verify_via_qr_url(driver, qr_url: str, row_data: dict) -> str:
    """
    QR 원본 URL(hometax.page.link)이나 내부 홈택스 URL을 Selenium으로 열어
    결과 페이지를 파싱하는 전략.

    홈택스 모바일 페이지 → "발급된 사실이 있습니다" 등 문구 확인.

    Parameters
    ----------
    driver  : Chrome WebDriver
    qr_url  : results.xlsx 의 원본QR 값
    row_data: 승인번호, 공급자번호 등 딕셔너리

    Returns
    -------
    'O' / 'X' / '' (판정 불가 시 빈 문자열 → 다음 전략으로)
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    # ── QR URL에서 내부 Hometax URL 추출 ─────────────────
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


def _extract_inner_url(qr_url: str) -> str:
    """
    hometax.page.link QR URL에서 실제 홈택스 접근 URL 추출.
    link= 파라미터의 이중 인코딩 해제 후 반환.
    """
    try:
        # link= 이후 부분 추출
        decoded = unquote(unquote(qr_url))
        if "link=" in decoded:
            inner = decoded.split("link=", 1)[1]
            # action= 파라미터가 있으면 그게 실제 API URL
            if "action=" in inner:
                action_part = inner.split("action=", 1)[1]
                return unquote(action_part.split("&")[0])
            return inner
        # 이미 홈택스 URL인 경우
        if "hometax.go.kr" in qr_url or "mob.tbet" in qr_url:
            return qr_url
    except Exception:
        pass
    return ""


# ══════════════════════════════════════════════
# 4. 전략 2: 모바일 API 직접 호출 (requests)
# ══════════════════════════════════════════════
def verify_via_api(row_data: dict) -> str:
    """
    홈택스 모바일 JSON API 엔드포인트를 직접 호출하여 진위 확인.

    호출 URL:
        https://mob.tbet.hometax.go.kr/jsonAction.do
        ?actionId=UTBETBDA16F001
        &etan=<승인번호>
        &splrTxprDscmNo=<공급자번호>
        &wrtDt=<작성일자>
        &splrCft=<합계금액>

    Returns
    -------
    'O' / 'X' / '' (응답 파싱 불가 시 빈 문자열)
    """
    import requests

    params = {
        "actionId":        "UTBETBDA16F001",
        "menuId":          "6001020100",
        "etan":            row_data["승인번호_원본"],   # 원본값 그대로 (cr 포함)
        "splrTxprDscmNo":  row_data["공급자번호"],
        "wrtDt":           row_data["작성일자_raw"],
        "splrCft":         row_data["합계금액"],
    }
    if row_data.get("수급자번호"):
        params["dmnrTxprDscmNo"] = row_data["수급자번호"]

    headers = {
        "User-Agent":    "Mozilla/5.0 (Linux; Android 13; SM-G991B) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/131.0.0.0 Mobile Safari/537.36",
        "Accept":        "application/json, text/plain, */*",
        "Referer":       "https://mob.tbet.hometax.go.kr/",
        "Origin":        "https://mob.tbet.hometax.go.kr",
    }

    log.info(f"  [전략2] 모바일 API 호출: {MOBILE_API}?{urlencode(params)[:80]}...")
    try:
        resp = requests.get(
            MOBILE_API, params=params, headers=headers,
            timeout=15, verify=False
        )
        log.info(f"  [전략2] 응답 코드: {resp.status_code}")
        log.debug(f"  [전략2] 응답 내용: {resp.text[:300]}")

        text = resp.text
        # JSON 응답 파싱
        try:
            data = resp.json()
            # 응답 구조에 따라 결과 추출 (실제 응답 구조 확인 후 조정 필요)
            result_code = (data.get("resultCode") or data.get("errCd") or
                           data.get("result") or "")
            result_msg  = (data.get("resultMsg") or data.get("errMsg") or
                           data.get("message") or "")
            text = str(result_code) + " " + str(result_msg)
        except Exception:
            pass   # JSON 파싱 실패 시 text 그대로 사용

        return _judge_text(text)

    except Exception as e:
        log.warning(f"  [전략2] API 호출 오류: {e}")
        return ""


# ══════════════════════════════════════════════
# 5. 전략 3: 홈택스 WebSquare 메뉴 JS 내비게이션
# ══════════════════════════════════════════════
def verify_via_websquare(driver, row_data: dict) -> str:
    """
    홈택스 메인 페이지에서 JavaScript를 통해 WebSquare 메뉴를 탐색하고
    전자세금계산서 발급사실 조회 폼을 자동 입력.

    WebSquare 특성:
        - URL은 메인 index로 고정됨 (SPA 구조)
        - 실제 콘텐츠는 mf_txppIframe에 동적 로드됨
        - 메뉴 이동: gfn_viewMenu(menuCd) 또는 특정 JS 함수 호출

    Returns
    -------
    'O' / 'X' / '확인불가'
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException

    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    log.info(f"  [전략3] 홈택스 WebSquare 내비게이션 시작")

    # ── 메인 페이지 로드 ──────────────────────────────
    driver.get(HOMETAX_MAIN)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(5)

    # ── JavaScript로 메뉴 이동 시도 ──────────────────
    nav_success = False
    for menu_cd in ETAX_MENU_CANDIDATES:
        try:
            # WebSquare 메뉴 이동 함수 (버전에 따라 다를 수 있음)
            driver.execute_script(f"gfn_viewMenu('{menu_cd}')")
            time.sleep(3)
            # iframe 내 페이지 변경 확인
            body_text = driver.find_element(By.TAG_NAME, "body").text
            if "전자세금계산서" in body_text or "승인번호" in body_text:
                log.info(f"  [전략3] 메뉴 이동 성공 (menuCd={menu_cd})")
                nav_success = True
                break
        except Exception:
            pass

    if not nav_success:
        # iframe 내부 탐색
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
        log.warning("  [전략3] 자동 메뉴 이동 실패 → 사용자 안내")
        return _guide_manual_navigation(driver, row_data)

    # ── 폼 입력 ──────────────────────────────────────
    result = _fill_websquare_form(driver, row_data)
    return result


def _fill_websquare_form(driver, row_data: dict) -> str:
    """WebSquare 폼에 데이터 입력 후 결과 판독"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import NoSuchElementException

    # 입력 필드 ID 후보 (홈택스 버전별 다름 → F12 개발자 도구로 확인 권장)
    field_map = {
        "승인번호":   (row_data["승인번호_정제"], ["etan", "tbx_etan", "aprvNo", "inp_etan"]),
        "공급자번호": (row_data["공급자번호"],    ["splrTxprDscmNo", "tbx_splrBizNo", "inp_splrBizNo"]),
        "작성일자":   (row_data["작성일자_raw"],  ["wrtDt", "tbx_wrtDt", "inp_wrtDt"]),
        "합계금액":   (row_data["합계금액"],      ["splrCft", "tbx_splrCft", "inp_totAmt"]),
    }

    def _try_fill(candidates: list, value: str) -> bool:
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

    # ── 조회 버튼 클릭 ───────────────────────────────
    _click_search_button(driver)
    time.sleep(3)

    return _judge_page_text(driver) or "확인불가"


def _click_search_button(driver) -> None:
    """'조회하기' 버튼 클릭 (텍스트 및 ID 패턴 탐색)"""
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


def _guide_manual_navigation(driver, row_data: dict) -> str:
    """자동 메뉴 이동 실패 시 사용자 수동 안내 후 결과 판독"""
    print("\n" + "="*60)
    print("  ⚠  자동 메뉴 이동에 실패했습니다.")
    print("  브라우저에서 직접 아래 경로로 이동해 주세요:")
    print("  홈택스 → 조회/발급 → 전자세금계산서 → 발급사실 조회")
    print(f"\n  입력할 정보:")
    print(f"    승인번호:   {row_data['승인번호_정제']}")
    print(f"    공급자번호: {row_data['공급자번호']}")
    print(f"    작성일자:   {row_data['작성일자_raw']}")
    print(f"    합계금액:   {row_data['합계금액']}")
    print("\n  조회 후 결과를 입력하세요:")
    print("  O=정상  X=불일치  S=건너뜀  Q=전체종료")
    print("="*60)
    ans = input("  결과 입력 → ").strip().upper()
    if ans == "Q":
        raise SystemExit("사용자가 종료를 요청했습니다.")
    if ans in ("O", "X"):
        return ans
    return "확인불가"


# ══════════════════════════════════════════════
# 6. 결과 판정 유틸리티
# ══════════════════════════════════════════════
SUCCESS_TEXTS = [
    "발급된 사실이 있습니다",
    "발급사실이 있습니다",
    "정상 발급",
    "발급 확인",
    "true",
    "0000",          # API 성공 코드
]
FAILURE_TEXTS = [
    "발급 사실이 없습니다",
    "발급사실이 없습니다",
    "일치하지 않습니다",
    "조회되지 않습니다",
    "해당 자료가 없습니다",
    "없습니다",
    "false",
    "9999",          # API 실패 코드 패턴
]
CAPTCHA_TEXTS = [
    "자동입력방지", "보안문자", "CAPTCHA",
    "captcha", "자동화", "확인코드",
]


def _judge_text(text: str) -> str:
    """텍스트 문자열만으로 O/X/'' 판정"""
    t = str(text)
    for p in CAPTCHA_TEXTS:
        if p.lower() in t.lower():
            return _handle_captcha_text()
    for p in SUCCESS_TEXTS:
        if p in t:
            return "O"
    for p in FAILURE_TEXTS:
        if p in t:
            return "X"
    return ""


def _judge_page_text(driver) -> str:
    """현재 페이지(+ 모든 iframe) 텍스트로 O/X 판정"""
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


def _handle_captcha_text() -> str:
    """CAPTCHA 감지 시 사용자 안내 후 수동 결과 입력"""
    print("\n" + "="*60)
    print("  ⚠  보안문자(CAPTCHA) 감지!")
    print("  브라우저에서 보안문자 입력 후 조회를 완료하세요.")
    print("  완료 후 결과를 입력하세요: O=정상  X=불일치  확인불가=모름")
    print("="*60)
    ans = input("  결과 입력 → ").strip().upper()
    if ans in ("O", "X"):
        return ans
    return "확인불가"


# ══════════════════════════════════════════════
# 7. 단건 조회 (3중 전략 통합)
# ══════════════════════════════════════════════
def verify_one(driver, row_series: pd.Series) -> str:
    """
    단건 전자세금계산서 조회.
    전략 1 → 2 → 3 순으로 시도하여 첫 번째 유효한 결과 반환.

    Parameters
    ----------
    driver     : Chrome WebDriver
    row_series : results.xlsx 단일 행 (Series)

    Returns
    -------
    'O' / 'X' / '확인불가'
    """
    # 데이터 정제
    row_data = {
        "승인번호_원본": str(row_series.get("승인번호", "")),
        "승인번호_정제": clean_approval_number(str(row_series.get("승인번호", ""))),
        "공급자번호":    clean_number(row_series.get("공급자번호", "")),
        "작성일자_raw":  clean_date(str(row_series.get("작성일자", ""))),
        "합계금액":      clean_number(row_series.get("합계금액", "")),
        "수급자번호":    clean_number(row_series.get("수급자번호", "")),
    }
    qr_url = str(row_series.get("원본QR", ""))

    log.info(f"  조회 데이터: 승인번호={row_data['승인번호_정제']} "
             f"공급자={row_data['공급자번호']} "
             f"일자={row_data['작성일자_raw']} "
             f"금액={row_data['합계금액']}")

    # ── 전략 1: QR URL 직접 접근 ───────────────────
    if qr_url and qr_url.startswith("http"):
        result = verify_via_qr_url(driver, qr_url, row_data)
        if result:
            return result

    # ── 전략 2: 모바일 API 직접 호출 ──────────────
    result = verify_via_api(row_data)
    if result:
        return result

    # ── 전략 3: WebSquare 메뉴 내비게이션 ─────────
    result = verify_via_websquare(driver, row_data)
    return result or "확인불가"


# ══════════════════════════════════════════════
# 8. 일괄 처리
# ══════════════════════════════════════════════
def run_batch(input_path: str, output_path: str) -> None:
    """
    results.xlsx의 전 건을 조회하고 verification_final.xlsx 저장.
    중간 저장으로 중단 재시작 시 이미 완료된 건은 건너뜁니다.
    """
    # ── 데이터 로드 ──────────────────────────────────
    df = pd.read_excel(input_path)
    log.info(f"로드: {len(df)}건 / 컬럼: {df.columns.tolist()}")

    required = ["승인번호", "공급자번호", "작성일자", "합계금액"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        log.error(f"필수 컬럼 누락: {missing}")
        sys.exit(1)

    if "진위여부" not in df.columns:
        df["진위여부"] = ""

    # ── 브라우저 시작 ────────────────────────────────
    log.info("Chrome 브라우저 시작 중...")
    driver = setup_driver()

    log.info("="*55)
    log.info(f"총 {len(df)}건 조회 시작")
    log.info("="*55)

    try:
        for idx, row in df.iterrows():
            existing = str(row.get("진위여부", "")).strip()
            if existing in ("O", "X", "확인불가"):
                log.info(f"[{idx+1}/{len(df)}] {row['파일명']}: 이미 처리됨({existing}), 건너뜀")
                continue

            log.info(f"[{idx+1}/{len(df)}] {row['파일명']}")

            result = verify_one(driver, row)
            df.at[idx, "진위여부"] = result
            log.info(f"  최종 판정: {result}")

            # 중간 저장
            _save_excel(df, output_path)
            log.info(f"  중간 저장 → {output_path}")

            # 랜덤 대기
            wait_sec = random.uniform(SLEEP_MIN, SLEEP_MAX)
            log.info(f"  다음 조회까지 {wait_sec:.1f}초 대기...")
            time.sleep(wait_sec)
            log.info("-"*55)

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # ── 최종 저장 및 요약 ────────────────────────────
    _save_excel(df, output_path)
    ok  = (df["진위여부"] == "O").sum()
    ng  = (df["진위여부"] == "X").sum()
    unk = (df["진위여부"] == "확인불가").sum()

    log.info("="*55)
    log.info(f"완료! → {output_path}")
    log.info(f"  O(정상): {ok}건  |  X(불일치): {ng}건  |  확인불가: {unk}건")
    log.info("="*55)


def _save_excel(df: pd.DataFrame, path: str) -> None:
    """DataFrame → Excel 저장 + 열 너비 자동 조정"""
    df.to_excel(path, index=False, engine="openpyxl")
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active
        for col in ws.columns:
            ml = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 60)
        wb.save(path)
    except Exception:
        pass


# ══════════════════════════════════════════════
# 9. 진입점
# ══════════════════════════════════════════════
if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings()   # SSL 경고 억제

    base        = Path(__file__).parent
    input_path  = str(base / INPUT_FILE)
    output_path = str(base / OUTPUT_FILE)

    if not Path(input_path).exists():
        log.error(f"입력 파일 없음: {input_path}")
        sys.exit(1)

    print("\n" + "="*60)
    print("  전자세금계산서 홈택스 진위 확인 자동화")
    print("="*60)
    print(f"  입력 : {input_path}")
    print(f"  출력 : {output_path}")
    print()
    print("  ※ 동작 전략 (자동 순차 시도)")
    print("    1단계: QR URL → 홈택스 모바일 웹 결과 접근")
    print("    2단계: 홈택스 모바일 API 직접 호출")
    print("    3단계: 홈택스 WebSquare 메뉴 + 폼 자동 입력")
    print()
    print("  ※ 주의사항")
    print("    - Chrome 창이 열립니다 (CAPTCHA 수동 입력 대비)")
    print("    - 보안문자 발생 시 터미널의 안내를 따르세요")
    print("    - 폼 ID가 안 맞는 경우 F12(개발자 도구)로 확인 후")
    print("      hometax_verify.py의 FIELD_IDS를 수정하세요")
    print("="*60 + "\n")

    run_batch(input_path, output_path)
