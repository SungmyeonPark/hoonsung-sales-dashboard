import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime, timedelta
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from dateutil.relativedelta import relativedelta
import time
import ssl
from googleapiclient.errors import HttpError

st.set_page_config(
    page_title="훈성 매출 분석 대시보드",
    layout="wide"
)

st.title("훈성 매출 분석 대시보드")


# --------------------------------------------------
# Keyword dictionaries
# --------------------------------------------------

CARD_SETTLEMENT_KEYWORDS = [
    "현대", "삼성", "롯데", "신한", "KB", "NH", "하나", "BC",
    "730674", "우600",
    "18918207", "9941183135", "0128433687",
    "10745627", "16505356", "91264283"
]

SEAFOOD_MEAT_KEYWORDS = [
    "생선값", "수산", "활어", "축산", "육류",
    "이병직", "연세활어", "활주로수산", "수원축산물",
    "주식회사은해수", "은해수"
]

GENERAL_FOOD_KEYWORDS = [
    "식자재", "대원식자재", "모노마트", "마트킹",
    "케이디", "희망수", "가보종", "장터", "마요네즈",
    "광천유통", "과일놀이터", "주식회사서림유",
    "금고주식회사서림유통", "서림유통",
    "대방할인마트", "(주)대방할인마트", "트레이더스수원", "트레이더스", "엔마트세류점", "엔마트"
]

ALCOHOL_KEYWORDS = [
    "주류구매", "수원주류", "주류판매", "주류"
]

SUPPLIES_KEYWORDS = [
    "아성다이소", "다이소"
]

RENT_KEYWORDS = [
    "기업(주)우성엔지니", "우성엔지니"
]

TAX_KEYWORDS = [
    "건강보험", "국민건강", "국민연금",
    "고용보험", "산재보험", "기금및기타국고",
    "수원팔달구", "국고",
    "수원시지방세", "수원시지방소", "지방세",
    "국세청", "중부지방국세", "중부지방국세청"
]

UTILITY_KEYWORDS = [
    "가스요금", "가스", "한전", "전기", "KT", "통신"
]

HYGIENE_KEYWORDS = [
    "세스코", "음식물협동", "루헨스", "정수기", "렌탈"
]

PROFESSIONAL_KEYWORDS = [
    "세무사", "이희석세무사"
]

FINANCE_INSURANCE_KEYWORDS = [
    "한국캐피탈", "메리츠", "캐피탈", "보험"
]

PLATFORM_FEE_KEYWORDS = [
    "윈글로벌페이", "쿠팡페이"
]

ONLINE_PURCHASE_KEYWORDS = [
    "쿠팡", "네이버페이", "와이에스인터내셔"
]

SMALL_PURCHASE_KEYWORDS = [
    "쿠팡이츠", "세븐일레븐", "씨유", "CU", "GS25"
]

KOREAN_SURNAMES = [
    "김", "이", "박", "최", "정", "강", "조", "윤", "장", "임",
    "한", "오", "서", "신", "권", "황", "안", "송", "전", "홍",
    "유", "고", "문", "양", "손", "배", "백", "허", "남", "심",
    "노", "하", "곽", "성", "차", "주", "우", "구", "민", "진",
    "류", "나", "엄", "원", "천", "방", "공", "현", "함", "변"
]

PERSON_NAME_EXCLUDE_KEYWORDS = [
    "주식회사", "(주)", "유한회사", "수원시", "지방세",
    "식자재", "마트", "유통", "주류", "수산", "축산",
    "가스", "한전", "KT", "세스코", "쿠팡", "네이버",
    "캐피탈", "보험", "건강", "국민건강", "고용보험", "산재보험",
    "대원", "모노", "마트킹", "서림", "대방", "다이소", "은해수"
]


# --------------------------------------------------
# Data preprocessing helpers
# --------------------------------------------------
@st.cache_data(ttl=3600, show_spinner="거래내역을 불러오고 분석하는 중입니다...")
def load_and_prepare_transactions_from_drive(folder_id):
    """
    Google Drive에서 거래내역 파일을 읽고,
    전체 전처리/분류까지 완료한 DataFrame을 캐싱합니다.

    이 함수가 핵심입니다.
    위젯 조작 때마다 Google Drive와 xls 파싱을 반복하지 않게 합니다.
    """

    drive_files = list_drive_transaction_files(folder_id)

    if not drive_files:
        return pd.DataFrame(), []

    all_dfs = []
    loaded_file_names = []

    for file_info in drive_files:
        file_name = file_info["name"]
        file_id = file_info["id"]
        modified_time = file_info.get("modifiedTime", "")

        file_bytes = download_drive_file_cached(
            file_id,
            file_name,
            modified_time
        )

        file_obj = make_file_like_object(file_bytes, file_name)

        temp_df = load_transaction_file(file_obj)
        temp_df = normalize_columns(temp_df)

        start_date, end_date = parse_woori_filename(file_name)

        temp_df["source_file"] = file_name
        temp_df["file_start_date"] = start_date
        temp_df["file_end_date"] = end_date

        all_dfs.append(temp_df)
        loaded_file_names.append(file_name)

    df = pd.concat(all_dfs, ignore_index=True)

    required_cols = [
        "거래일시",
        "기재내용",
        "찾으신금액",
        "맡기신금액",
        "거래후 잔액"
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        raise ValueError(f"필수 컬럼이 없습니다: {missing_cols}. 현재 컬럼명: {list(df.columns)}")

    df["거래일시"] = pd.to_datetime(df["거래일시"], errors="coerce")

    money_cols = ["찾으신금액", "맡기신금액", "거래후 잔액"]

    for col in money_cols:
        df[col] = clean_money_column(df[col])

    df = df.dropna(subset=["거래일시"])

    df["거래일"] = df["거래일시"].dt.date
    df["월"] = df["거래일시"].dt.to_period("M").astype(str)
    df["분기"] = df["거래일시"].dt.to_period("Q").astype(str)

    df["입금액"] = df["맡기신금액"]
    df["출금액"] = df["찾으신금액"]

    df["거래구분"] = df.apply(
        lambda row: "입금" if row["입금액"] > 0 else "출금",
        axis=1
    )

    df[["대분류", "세부분류"]] = df.apply(
        lambda row: pd.Series(classify_transaction(row)),
        axis=1
    )

    df["카테고리"] = df["대분류"] + " - " + df["세부분류"]

    df["영업매출"] = df.apply(
        lambda row: row["입금액"] if row["대분류"] == "매출/입금" else 0,
        axis=1
    )

    df["비영업입금"] = df.apply(
        lambda row: row["입금액"] if row["대분류"] == "비영업입금" else 0,
        axis=1
    )

    df["비용"] = df["출금액"]
    df["순현금흐름"] = df["입금액"] - df["출금액"]

    df = make_korean_weekday(df)
    df = df.sort_values("거래일시", ascending=True)

    duplicate_check_cols = [
        "거래일시",
        "적요",
        "기재내용",
        "찾으신금액",
        "맡기신금액",
        "거래후 잔액"
    ]

    duplicate_check_cols = [col for col in duplicate_check_cols if col in df.columns]

    df = df.drop_duplicates(subset=duplicate_check_cols, keep="first")

    return df, loaded_file_names
# --------------------------------------------------
# Google Drive helpers
# --------------------------------------------------

@st.cache_resource
def get_drive_service():
    """
    Create Google Drive API service using Streamlit secrets.
    """

    service_account_info = dict(st.secrets["gcp_service_account"])

    # private_key 줄바꿈 문제 방지
    if "private_key" in service_account_info:
        service_account_info["private_key"] = service_account_info["private_key"].replace("\\n", "\n")

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )

    return build("drive", "v3", credentials=credentials)


@st.cache_data(ttl=600, show_spinner="Google Drive 파일 목록을 확인하는 중입니다...")
def list_drive_transaction_files(folder_id):
    """
    Google Drive 폴더의 파일 목록을 캐싱해서 반복 조회를 줄입니다.
    ttl=600이면 10분 동안 같은 폴더 목록을 재사용합니다.
    """

    service = get_drive_service()

    query = f"'{folder_id}' in parents and trashed = false"

    files = []
    page_token = None

    while True:
        request = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, modifiedTime, size)",
            pageToken=page_token,
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        )

        results = execute_with_retry(request)

        files.extend(results.get("files", []))

        page_token = results.get("nextPageToken")

        if not page_token:
            break

    allowed_files = [
        file_info for file_info in files
        if file_info["name"].lower().endswith((".xls", ".xlsx", ".csv"))
    ]

    allowed_files = sorted(allowed_files, key=lambda x: x["name"])

    return allowed_files




@st.cache_data(ttl=3600, max_entries=100, show_spinner="Google Drive 파일을 다운로드하는 중입니다...")
def download_drive_file_cached(file_id, file_name, modified_time):
    """
    파일 ID + 파일명 + 수정시간을 기준으로 다운로드 결과를 캐싱합니다.
    파일이 바뀌면 modified_time이 바뀌므로 자동으로 다시 다운로드됩니다.
    """

    service = get_drive_service()

    request = service.files().get_media(fileId=file_id)
    file_buffer = BytesIO()

    downloader = MediaIoBaseDownload(file_buffer, request)

    done = False

    while not done:
        status, done = downloader.next_chunk()

    return file_buffer.getvalue()


def make_file_like_object(file_bytes, file_name):
    """
    bytes를 pandas가 읽을 수 있는 file-like object로 변환합니다.
    """

    file_buffer = BytesIO(file_bytes)
    file_buffer.seek(0)
    file_buffer.name = file_name

    return file_buffer


def load_files_from_google_drive():
    """
    Load all transaction files from Google Drive folder.
    """

    folder_id = st.secrets["DRIVE_FOLDER_ID"]

    drive_files = list_drive_transaction_files(folder_id)

    if not drive_files:
        return []

    downloaded_files = []

    for file_info in drive_files:
        downloaded_file = download_drive_file(
            file_info["id"],
            file_info["name"]
        )

        downloaded_files.append(downloaded_file)

    return downloaded_files


def parse_woori_filename(file_name):
    """
    Parse file name like 250508_260508.xls.
    Meaning: 2025-05-08 to 2026-05-08.
    """

    match = re.search(r"(\d{6})_(\d{6})", file_name)

    if not match:
        return None, None

    start_raw = match.group(1)
    end_raw = match.group(2)

    start_date = datetime.strptime(start_raw, "%y%m%d").date()
    end_date = datetime.strptime(end_raw, "%y%m%d").date()

    return start_date, end_date


def load_transaction_file(uploaded_file):
    """
    Load Woori Bank transaction file.
    Supports csv, xlsx, xls.

    Some Korean bank .xls files are actually HTML tables,
    so we try read_excel first and fall back to read_html.
    """

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="cp949")

    if file_name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, engine="openpyxl")

    if file_name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded_file, engine="xlrd")
        except Exception:
            uploaded_file.seek(0)
            tables = pd.read_html(uploaded_file)
            return tables[0]

    raise ValueError("지원하지 않는 파일 형식입니다.")


def normalize_columns(df):
    """
    Clean column names.
    If the actual header row was read as data, detect and fix it.
    """

    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    df.columns = df.columns.astype(str).str.strip()

    required_cols = ["거래일시", "기재내용", "찾으신금액", "맡기신금액", "거래후 잔액"]

    if all(col in df.columns for col in required_cols):
        return df

    for idx in range(min(10, len(df))):
        row_values = df.iloc[idx].astype(str).str.strip().tolist()

        if "거래일시" in row_values and "찾으신금액" in row_values and "맡기신금액" in row_values:
            new_columns = row_values
            df = df.iloc[idx + 1:].copy()
            df.columns = new_columns
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(how="all")
            df = df.dropna(axis=1, how="all")
            return df

    return df


def clean_money_column(series):
    cleaned = (
        series
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
        .replace(["", "nan", "None", "-", "NaN"], "0")
    )

    cleaned = cleaned.str.replace(r"[^0-9.-]", "", regex=True)
    cleaned = cleaned.replace(["", "-", ".", "-."], "0")

    return cleaned.astype(float)

def execute_with_retry(request, max_retries=3, sleep_seconds=2):
    """
    Google API 요청이 SSL/network 문제로 실패할 때 몇 번 재시도합니다.
    """

    for attempt in range(max_retries):
        try:
            return request.execute()
        except (ssl.SSLError, HttpError) as e:
            if attempt == max_retries - 1:
                raise e

            time.sleep(sleep_seconds * (attempt + 1))


# --------------------------------------------------
# Classification helpers
# --------------------------------------------------

def contains_any(text, keywords):
    return any(keyword in text for keyword in keywords)


def normalize_text(row):
    desc = str(row.get("기재내용", ""))
    summary = str(row.get("적요", ""))
    agency = str(row.get("취급기관", ""))

    return f"{desc} {summary} {agency}".replace(" ", "")


def has_korean_person_name(text):
    """
    Detect Korean personal names in transaction text.
    This is intentionally broad because outgoing transactions
    with Korean names should be treated as labor cost.
    """

    compact_text = str(text).replace(" ", "")

    if contains_any(compact_text, PERSON_NAME_EXCLUDE_KEYWORDS):
        return False

    # Case 1: bank name + Korean name
    # Examples: 기업김호재, 농협조재원, 국민박가연, 신한백승민, 토뱅윤상혁
    bank_name_pattern = r"(국민|농협|신한|우리|하나|기업|토뱅|카뱅|카카오|케이뱅크|새마을|우체국)([가-힣]{2,4})"
    bank_match = re.search(bank_name_pattern, compact_text)

    if bank_match:
        possible_name = bank_match.group(2)
        if possible_name and possible_name[0] in KOREAN_SURNAMES:
            return True

    # Case 2: exact Korean name only
    # Examples: 김지훈, 임윤정, 강승헌
    exact_name_pattern = r"^[가-힣]{2,4}$"

    if re.match(exact_name_pattern, compact_text):
        if compact_text[0] in KOREAN_SURNAMES:
            return True

    # Case 3: Korean name followed by memo mark
    # Examples: 김지훈(4월, 신한권지민(일일알바
    embedded_name_pattern = r"([가-힣]{2,4})(?:\(|-|$)"
    embedded_matches = re.findall(embedded_name_pattern, compact_text)

    for possible_name in embedded_matches:
        if possible_name and possible_name[0] in KOREAN_SURNAMES:
            return True

    return False


def looks_like_staff_transfer(row, text):
    """
    User rule:
    If outgoing transaction contains a Korean person name,
    treat it as labor cost.
    """

    expense = row.get("출금액", 0)

    if expense <= 0:
        return False

    if contains_any(text, ["알바", "일일알바", "알바인센", "급여", "인건비"]):
        return True

    if has_korean_person_name(text):
        return True

    return False


def classify_transaction(row):
    """
    Restaurant-focused transaction classifier.
    Returns:
    - 대분류
    - 세부분류
    """

    text = normalize_text(row)
    income = row.get("입금액", 0)
    summary = str(row.get("적요", ""))

    # -------------------------
    # Income classification
    # -------------------------

    if income > 0:
        if contains_any(text, ["체크카드취소", "취소", "환불"]):
            return "비영업입금", "환불/취소 입금"

        if contains_any(text, ["건강보험환급", "장기요양환급", "과오납연금", "환급"]):
            return "비영업입금", "보험/세금 환급"

        if contains_any(text, ["식자재", "생선값", "운영자금", "정산"]):
            return "비영업입금", "내부자금/정산 입금"

        if contains_any(text, ["쿠팡페이", "쿠페이", "네이버페이", "네이버페이지원"]):
            return "매출/입금", "플랫폼/간편결제 정산"

        if contains_any(text, CARD_SETTLEMENT_KEYWORDS):
            return "매출/입금", "카드매출 정산"

        if contains_any(summary, ["모바일", "타행건별", "펌뱅킹"]):
            return "매출/입금", "현금/계좌이체 매출"

        return "매출/입금", "기타 입금"

    # -------------------------
    # Expense classification
    # -------------------------

    if contains_any(text, RENT_KEYWORDS):
        return "임대료", "월세"

    if looks_like_staff_transfer(row, text):
        return "인건비", "급여/알바비"

    if contains_any(text, ALCOHOL_KEYWORDS):
        return "매출원가", "주류 매입"

    if contains_any(text, SEAFOOD_MEAT_KEYWORDS):
        return "매출원가", "수산/육류 매입"

    if contains_any(text, GENERAL_FOOD_KEYWORDS):
        return "매출원가", "일반 식자재/소모품"

    if contains_any(text, SUPPLIES_KEYWORDS):
        return "운영비", "소모품"

    if contains_any(text, ONLINE_PURCHASE_KEYWORDS):
        return "운영비", "온라인 구매/소모품"

    if contains_any(text, SMALL_PURCHASE_KEYWORDS):
        return "운영비", "식비/간식/소액구매"

    if contains_any(text, UTILITY_KEYWORDS):
        return "공과금", "가스/전기/통신"

    if contains_any(text, HYGIENE_KEYWORDS):
        return "위생/관리비", "방역/폐기물/렌탈"

    if contains_any(text, TAX_KEYWORDS):
        return "세금/보험", "세금/4대보험"

    if contains_any(text, PROFESSIONAL_KEYWORDS):
        return "전문서비스", "세무/회계"

    if contains_any(text, FINANCE_INSURANCE_KEYWORDS):
        return "금융/보험", "대출/보험료"

    if contains_any(text, PLATFORM_FEE_KEYWORDS):
        return "수수료", "결제/플랫폼 수수료"

    return "기타", "미분류"


# --------------------------------------------------
# Dashboard helpers
# --------------------------------------------------

def to_date_value(x):
    """
    Convert pandas Timestamp or datetime to Python date.
    """

    if isinstance(x, pd.Timestamp):
        return x.date()

    if isinstance(x, datetime):
        return x.date()

    return x


def date_range_selector(label, key, min_date, max_date, default_start=None, default_end=None):
    """
    Reusable date range selector.
    Direct input uses one date range picker instead of two separate date inputs.
    This avoids disabled-year issues caused by Streamlit widget state.
    """

    min_date = to_date_value(min_date)
    max_date = to_date_value(max_date)

    if default_start is None:
        default_start = min_date

    if default_end is None:
        default_end = max_date

    default_start = to_date_value(default_start)
    default_end = to_date_value(default_end)

    # Clamp defaults inside valid range
    default_start = max(min_date, min(default_start, max_date))
    default_end = max(min_date, min(default_end, max_date))

    if default_start > default_end:
        default_start, default_end = default_end, default_start

    st.markdown(f"**{label} 기간 선택**")

    mode = st.radio(
        f"{label} 기간 선택 방식",
        ["직접 입력", "슬라이더"],
        horizontal=True,
        key=f"{key}_mode"
    )

    if mode == "직접 입력":
        selected_range = st.date_input(
            "기간",
            value=(default_start, default_end),
            min_value=min_date,
            max_value=max_date,
            key=f"{key}_date_range"
        )

        if isinstance(selected_range, tuple):
            if len(selected_range) == 2:
                start_date, end_date = selected_range
            elif len(selected_range) == 1:
                start_date = selected_range[0]
                end_date = selected_range[0]
            else:
                start_date, end_date = default_start, default_end
        else:
            start_date = selected_range
            end_date = selected_range

    else:
        start_date, end_date = st.slider(
            "기간 슬라이더",
            min_value=min_date,
            max_value=max_date,
            value=(default_start, default_end),
            format="YYYY-MM-DD",
            key=f"{key}_slider"
        )

    if start_date > end_date:
        st.warning("시작일이 종료일보다 늦어 날짜를 자동으로 바꿨습니다.")
        start_date, end_date = end_date, start_date

    return start_date, end_date


def filter_by_date(df, start_date, end_date):
    return df[
        (df["거래일"] >= start_date) &
        (df["거래일"] <= end_date)
    ].copy()

def format_short_date(d):
    return d.strftime("%y-%m-%d")


def format_comparison_label(name, start_date, end_date):
    return f"{name}<br>{format_short_date(start_date)} ~ {format_short_date(end_date)}"


def get_month_end(d):
    next_month = d.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)


def build_recent_month_options(min_date, max_date):
    """
    Build monthly dropdown options within the recent 1-year window.
    Most recent month appears first.
    """

    start_limit = max(min_date, max_date - timedelta(days=365))

    month_starts = pd.date_range(
        start=start_limit.replace(day=1),
        end=max_date.replace(day=1),
        freq="MS"
    )

    options = []

    for month_start_ts in month_starts:
        month_start = month_start_ts.date()
        month_end = get_month_end(month_start)

        actual_start = max(month_start, min_date)
        actual_end = min(month_end, max_date)

        if actual_start <= actual_end:
            options.append({
                "label": f"{actual_start.strftime('%Y년 %m월')} ({format_short_date(actual_start)} ~ {format_short_date(actual_end)})",
                "start": actual_start,
                "end": actual_end
            })

    return list(reversed(options))


def build_recent_week_options(min_date, max_date):
    """
    Build weekly dropdown options within the recent 1-year window.
    Week is Monday to Sunday.
    Most recent week appears first.
    """

    start_limit = max(min_date, max_date - timedelta(days=365))

    # Monday of the max_date week
    current_week_start = max_date - timedelta(days=max_date.weekday())

    options = []

    week_start = current_week_start

    while week_start >= start_limit - timedelta(days=6):
        week_end = week_start + timedelta(days=6)

        actual_start = max(week_start, start_limit, min_date)
        actual_end = min(week_end, max_date)

        if actual_start <= actual_end:
            options.append({
                "label": f"{format_short_date(actual_start)} ~ {format_short_date(actual_end)}",
                "start": actual_start,
                "end": actual_end
            })

        week_start = week_start - timedelta(days=7)

    return options


def summarize_period(df, start_date, end_date, label):
    period_df = filter_by_date(df, start_date, end_date)

    return {
        "비교기간": label,
        "시작일": start_date,
        "종료일": end_date,
        "영업매출": period_df["영업매출"].sum(),
        "비용": period_df["비용"].sum(),
        "순현금흐름": period_df["순현금흐름"].sum(),
        "거래건수": len(period_df)
    }


def calculate_change_rate(current_value, comparison_value):
    if comparison_value == 0:
        return None

    return (current_value - comparison_value) / comparison_value

def summarize_period(df, start_date, end_date, label):
    """
    Summarize sales, expense, and cash flow for a selected period.
    """

    period_df = filter_by_date(df, start_date, end_date)

    return {
        "비교기간": label,
        "시작일": start_date,
        "종료일": end_date,
        "영업매출": period_df["영업매출"].sum(),
        "비용": period_df["비용"].sum(),
        "순현금흐름": period_df["순현금흐름"].sum(),
        "거래건수": len(period_df)
    }


def calculate_change_rate(current_value, comparison_value):
    """
    Calculate percentage change.
    """

    if comparison_value == 0:
        return None

    return (current_value - comparison_value) / comparison_value

def make_korean_weekday(df):
    weekday_map = {
        0: "월요일",
        1: "화요일",
        2: "수요일",
        3: "목요일",
        4: "금요일",
        5: "토요일",
        6: "일요일"
    }

    df["요일번호"] = df["거래일시"].dt.weekday
    df["요일"] = df["요일번호"].map(weekday_map)

    return df


# --------------------------------------------------
# Main app
# --------------------------------------------------

# --------------------------------------------------
# Main app
# --------------------------------------------------

try:
    folder_id = st.secrets["DRIVE_FOLDER_ID"].strip()

    st.info("Google Drive 폴더에서 거래내역 파일을 자동으로 불러옵니다.")

    refresh_cache = st.button("데이터 새로고침")

    if refresh_cache:
        st.cache_data.clear()
        st.rerun()

    df, loaded_file_names = load_and_prepare_transactions_from_drive(folder_id)

    if df.empty:
        st.warning("Google Drive 폴더에 .xls, .xlsx, .csv 거래내역 파일이 없습니다.")
        st.stop()

    st.success(f"Google Drive에서 {len(loaded_file_names)}개 파일을 불러왔습니다.")

    with st.expander("불러온 파일 목록 보기"):
        st.write(loaded_file_names)

    # --------------------------------------------------
    # Sidebar filters
    # --------------------------------------------------

    st.sidebar.header("전체 필터")

    available_types = sorted(df["거래구분"].dropna().unique())
    selected_types = st.sidebar.multiselect(
        "거래구분 선택",
        available_types,
        default=available_types
    )

    available_main_categories = sorted(df["대분류"].dropna().unique())
    selected_main_categories = st.sidebar.multiselect(
        "대분류 선택",
        available_main_categories,
        default=available_main_categories
    )

    available_sub_categories = sorted(df["세부분류"].dropna().unique())
    selected_sub_categories = st.sidebar.multiselect(
        "세부분류 선택",
        available_sub_categories,
        default=available_sub_categories
    )

    available_files = sorted(df["source_file"].dropna().unique())
    selected_files = st.sidebar.multiselect(
        "파일 선택",
        available_files,
        default=available_files
    )

    base_df = df[
        (df["거래구분"].isin(selected_types)) &
        (df["대분류"].isin(selected_main_categories)) &
        (df["세부분류"].isin(selected_sub_categories)) &
        (df["source_file"].isin(selected_files))
    ].copy()

    if base_df.empty:
        st.warning("선택된 필터에 해당하는 거래내역이 없습니다.")
        st.stop()

    data_min_date = base_df["거래일"].min()
    data_max_date = base_df["거래일"].max()

    st.caption(
        f"현재 선택된 총 거래 건수: {len(base_df):,}건 | "
        f"전체 기간: {data_min_date} ~ {data_max_date}"
    )

    # --------------------------------------------------
    # 1. Summary
    # --------------------------------------------------

    st.divider()
    st.header("1. 주요 요약 및 추이")

    current_year_start = data_max_date.replace(month=1, day=1)

    default_summary_start = max(data_min_date, current_year_start)
    default_summary_end = data_max_date

    summary_start, summary_end = date_range_selector(
        "주요 요약",
        "summary",
        data_min_date,
        data_max_date,
        default_summary_start,
        default_summary_end
    )

    summary_df = filter_by_date(base_df, summary_start, summary_end)

    total_income = summary_df["입금액"].sum()
    operating_sales = summary_df["영업매출"].sum()
    non_operating_income = summary_df["비영업입금"].sum()
    total_expense = summary_df["비용"].sum()
    net_cashflow = summary_df["순현금흐름"].sum()
    transaction_count = len(summary_df)

    seafood_meat_cost = summary_df[
        summary_df["세부분류"] == "수산/육류 매입"
    ]["비용"].sum()

    general_food_cost = summary_df[
        summary_df["세부분류"] == "일반 식자재/소모품"
    ]["비용"].sum()

    alcohol_cost = summary_df[
        summary_df["세부분류"] == "주류 매입"
    ]["비용"].sum()

    total_food_alcohol_cost = seafood_meat_cost + general_food_cost + alcohol_cost

    labor_cost = summary_df[
        summary_df["대분류"] == "인건비"
    ]["비용"].sum()

    rent_cost = summary_df[
        summary_df["대분류"] == "임대료"
    ]["비용"].sum()

    utility_cost = summary_df[
        summary_df["대분류"].isin(["공과금", "위생/관리비"])
    ]["비용"].sum()

    total_food_alcohol_ratio = (
        total_food_alcohol_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    seafood_meat_ratio = (
        seafood_meat_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    general_food_ratio = (
        general_food_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    alcohol_ratio = (
        alcohol_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    labor_cost_ratio = (
        labor_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    rent_ratio = (
        rent_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    utility_ratio = (
        utility_cost / operating_sales
        if operating_sales > 0
        else 0
    )

    st.subheader("요약 지표")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("영업매출", f"{operating_sales:,.0f}원")
    c2.metric("총 비용", f"{total_expense:,.0f}원")
    c3.metric("순현금흐름", f"{net_cashflow:,.0f}원")
    c4.metric("거래 건수", f"{transaction_count:,}건")

    c5, c6, c7, c8 = st.columns(4)

    c5.metric("총 입금액", f"{total_income:,.0f}원")
    c6.metric("비영업입금", f"{non_operating_income:,.0f}원")
    c7.metric("인건비", f"{labor_cost:,.0f}원")
    c8.metric("임대료", f"{rent_cost:,.0f}원")

    st.subheader("식재료/주류 비율")

    r1, r2, r3, r4 = st.columns(4)

    r1.metric("식자재/주류 합산 비율", f"{total_food_alcohol_ratio:.1%}")
    r2.metric("수산/육류 비율", f"{seafood_meat_ratio:.1%}")
    r3.metric("일반 식자재/소모품 비율", f"{general_food_ratio:.1%}")
    r4.metric("주류 비율", f"{alcohol_ratio:.1%}")

    st.subheader("운영 핵심 비율")

    o1, o2, o3 = st.columns(3)

    o1.metric("인건비율", f"{labor_cost_ratio:.1%}")
    o2.metric("임대료율", f"{rent_ratio:.1%}")
    o3.metric("공과/관리비율", f"{utility_ratio:.1%}")

    # --------------------------------------------------
    # 2. Monthly / quarterly chart
    # --------------------------------------------------

    st.divider()
    st.header("2. 월별 / 분기별 영업매출, 비용, 순현금흐름")

    default_period_start = max(data_min_date, data_max_date - timedelta(days=365))
    default_period_end = data_max_date

    period_start, period_end = date_range_selector(
        "월별/분기별 차트",
        "period_chart",
        data_min_date,
        data_max_date,
        default_period_start,
        default_period_end
    )

    period_df = filter_by_date(base_df, period_start, period_end)

    monthly_tab, quarterly_tab = st.tabs(["월별", "분기별"])

    with monthly_tab:
        monthly_chart = (
            period_df
            .groupby("월")[["영업매출", "비용", "순현금흐름"]]
            .sum()
            .reset_index()
        )

        if not monthly_chart.empty:
            fig_monthly = px.bar(
                monthly_chart,
                x="월",
                y=["영업매출", "비용", "순현금흐름"],
                barmode="group",
                labels={
                    "월": "월",
                    "value": "금액",
                    "variable": "항목"
                },
                color_discrete_map={
                    "영업매출": "skyblue",
                    "비용": "red",
                    "순현금흐름": "green"
                }
            )
            st.plotly_chart(fig_monthly, use_container_width=True)
        else:
            st.info("선택된 기간에 해당하는 월별 데이터가 없습니다.")

    with quarterly_tab:
        quarterly_chart = (
            period_df
            .groupby("분기")[["영업매출", "비용", "순현금흐름"]]
            .sum()
            .reset_index()
        )

        if not quarterly_chart.empty:
            fig_quarterly = px.bar(
                quarterly_chart,
                x="분기",
                y=["영업매출", "비용", "순현금흐름"],
                barmode="group",
                labels={
                    "분기": "분기",
                    "value": "금액",
                    "variable": "항목"
                },
                color_discrete_map={
                    "영업매출": "skyblue",
                    "비용": "red",
                    "순현금흐름": "green"
                }
            )
            st.plotly_chart(fig_quarterly, use_container_width=True)
        else:
            st.info("선택된 기간에 해당하는 분기별 데이터가 없습니다.")

    # --------------------------------------------------
    # 3. Previous month / previous year comparison
    # --------------------------------------------------

    st.divider()
    st.header("3. 전월 / 전년 동월 비교")

    comparison_type = st.radio(
        "비교 기간 선택 방식",
        ["월별", "주별", "직접 입력"],
        horizontal=True,
        key="comparison_type"
    )

    comparison_start = None
    comparison_end = None

    if comparison_type == "월별":
        month_options = build_recent_month_options(data_min_date, data_max_date)

        if not month_options:
            st.warning("선택 가능한 월별 기간이 없습니다.")
            st.stop()

        selected_period = st.selectbox(
            "최근 1년 월 선택",
            month_options,
            format_func=lambda x: x["label"],
            key="comparison_month_select"
        )

        comparison_start = selected_period["start"]
        comparison_end = selected_period["end"]

    elif comparison_type == "주별":
        week_options = build_recent_week_options(data_min_date, data_max_date)

        if not week_options:
            st.warning("선택 가능한 주별 기간이 없습니다.")
            st.stop()

        selected_period = st.selectbox(
            "최근 1년 주 선택",
            week_options,
            format_func=lambda x: x["label"],
            key="comparison_week_select"
        )

        comparison_start = selected_period["start"]
        comparison_end = selected_period["end"]

    else:
        d1, d2 = st.columns(2)

        with d1:
            comparison_start = st.date_input(
                "비교 시작일",
                value=max(data_min_date, data_max_date - timedelta(days=30)),
                key="comparison_direct_start"
            )

        with d2:
            comparison_end = st.date_input(
                "비교 종료일",
                value=data_max_date,
                key="comparison_direct_end"
            )

        # 직접 입력은 min/max를 date_input에 직접 걸지 않고,
        # 입력 후 범위를 보정합니다. 그래야 연도 비활성화 문제가 줄어듭니다.
        if comparison_start < data_min_date:
            comparison_start = data_min_date

        if comparison_end > data_max_date:
            comparison_end = data_max_date

    if comparison_start is None or comparison_end is None:
        st.warning("비교 기간을 선택할 수 없습니다.")
        st.stop()

    if comparison_start > comparison_end:
        st.warning("시작일이 종료일보다 늦어 날짜를 자동으로 바꿨습니다.")
        comparison_start, comparison_end = comparison_end, comparison_start

    previous_month_start = comparison_start - relativedelta(months=1)
    previous_month_end = comparison_end - relativedelta(months=1)

    previous_year_start = comparison_start - relativedelta(years=1)
    previous_year_end = comparison_end - relativedelta(years=1)

    current_label = format_comparison_label(
        "기준",
        comparison_start,
        comparison_end
    )

    previous_month_label = format_comparison_label(
        "전월",
        previous_month_start,
        previous_month_end
    )

    previous_year_label = format_comparison_label(
        "전년 동월",
        previous_year_start,
        previous_year_end
    )

    current_summary = summarize_period(
        base_df,
        comparison_start,
        comparison_end,
        current_label
    )

    previous_month_summary = summarize_period(
        base_df,
        previous_month_start,
        previous_month_end,
        previous_month_label
    )

    previous_year_summary = summarize_period(
        base_df,
        previous_year_start,
        previous_year_end,
        previous_year_label
    )

    comparison_df = pd.DataFrame([
        current_summary,
        previous_month_summary,
        previous_year_summary
    ])

    comparison_metric = st.radio(
        "비교할 지표",
        ["영업매출", "비용", "순현금흐름"],
        horizontal=True,
        key="comparison_metric"
    )

    comparison_color_map = {
        current_label: "skyblue",
        previous_month_label: "orange",
        previous_year_label: "lightgreen"
    }

    fig_comparison = px.bar(
        comparison_df,
        x="비교기간",
        y=comparison_metric,
        text=comparison_metric,
        labels={
            "비교기간": "",
            comparison_metric: "금액"
        },
        color="비교기간",
        color_discrete_map=comparison_color_map
    )

    fig_comparison.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside"
    )

    fig_comparison.update_layout(
        showlegend=False,
        yaxis_title="금액",
        xaxis_title=None,
        xaxis_tickangle=0
    )

    st.plotly_chart(fig_comparison, use_container_width=True)

    current_value = current_summary[comparison_metric]
    previous_month_value = previous_month_summary[comparison_metric]
    previous_year_value = previous_year_summary[comparison_metric]

    mom_change = calculate_change_rate(current_value, previous_month_value)
    yoy_change = calculate_change_rate(current_value, previous_year_value)

    st.subheader("증감률 요약")

    m1, m2, m3 = st.columns(3)

    m1.metric(
        "기준",
        f"{current_value:,.0f}원"
    )

    if mom_change is None:
        m2.metric(
            "전월 대비",
            "비교 불가",
            delta="전월 데이터 없음"
        )
    else:
        m2.metric(
            "전월 대비",
            f"{current_value - previous_month_value:,.0f}원",
            delta=f"{mom_change:.1%}"
        )

    if yoy_change is None:
        m3.metric(
            "전년 동월 대비",
            "비교 불가",
            delta="전년 데이터 없음"
        )
    else:
        m3.metric(
            "전년 동월 대비",
            f"{current_value - previous_year_value:,.0f}원",
            delta=f"{yoy_change:.1%}"
        )

    st.subheader("비교 기간 상세")

    comparison_display_df = comparison_df.copy()
    comparison_display_df["비교기간"] = comparison_display_df["비교기간"].str.replace("<br>", "\n", regex=False)

    for col in ["영업매출", "비용", "순현금흐름"]:
        comparison_display_df[col] = comparison_display_df[col].map(lambda x: f"{x:,.0f}원")

    st.dataframe(
        comparison_display_df,
        use_container_width=True
    )

    # --------------------------------------------------
    # 4. Weekday chart
    # --------------------------------------------------

    st.divider()
    st.header("4. 요일별 매출 / 비용 분석")

    weekday_start, weekday_end = date_range_selector(
        "요일별 분석",
        "weekday_chart",
        data_min_date,
        data_max_date,
        data_min_date,
        data_max_date
    )

    weekday_df = filter_by_date(base_df, weekday_start, weekday_end)

    weekday_metric = st.radio(
        "요일별 집계 방식",
        ["총액", "평균"],
        horizontal=True,
        key="weekday_metric"
    )

    weekday_order = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]

    if weekday_metric == "총액":
        weekday_chart = (
            weekday_df
            .groupby(["요일번호", "요일"])[["영업매출", "비용"]]
            .sum()
            .reset_index()
            .sort_values("요일번호")
        )
    else:
        daily_weekday = (
            weekday_df
            .groupby(["거래일", "요일번호", "요일"])[["영업매출", "비용"]]
            .sum()
            .reset_index()
        )

        weekday_chart = (
            daily_weekday
            .groupby(["요일번호", "요일"])[["영업매출", "비용"]]
            .mean()
            .reset_index()
            .sort_values("요일번호")
        )

    if not weekday_chart.empty:
        fig_weekday = px.bar(
            weekday_chart,
            x="요일",
            y=["영업매출", "비용"],
            barmode="group",
            category_orders={"요일": weekday_order},
            labels={
                "요일": "요일",
                "value": "금액",
                "variable": "항목"
            },
            color_discrete_map={
                "영업매출": "skyblue",
                "비용": "red"
            }
        )
        st.plotly_chart(fig_weekday, use_container_width=True)
    else:
        st.info("선택된 기간에 해당하는 요일별 데이터가 없습니다.")

    # --------------------------------------------------
    # 5. Daily trend
    # --------------------------------------------------

    st.divider()
    st.header("5. 일별 영업매출 / 비용 추이")

    daily_min_date = max(data_min_date, data_max_date - timedelta(days=365))

    default_daily_start = max(daily_min_date, data_max_date - timedelta(days=14))
    default_daily_end = data_max_date

    daily_start, daily_end = date_range_selector(
        "일별 추이",
        "daily_trend",
        daily_min_date,
        data_max_date,
        default_daily_start,
        default_daily_end
    )

    daily_df = filter_by_date(base_df, daily_start, daily_end)

    show_sales = st.checkbox("영업매출 보기", value=True, key="show_daily_sales")
    show_expense = st.checkbox("비용 보기", value=True, key="show_daily_expense")

    daily_cols = []

    if show_sales:
        daily_cols.append("영업매출")

    if show_expense:
        daily_cols.append("비용")

    daily_trend = (
        daily_df
        .groupby("거래일")[["영업매출", "비용"]]
        .sum()
        .reset_index()
    )

    if not daily_cols:
        st.info("표시할 항목을 하나 이상 선택해주세요.")
    elif not daily_trend.empty:
        fig_daily = px.line(
            daily_trend,
            x="거래일",
            y=daily_cols,
            markers=True,
            labels={
                "거래일": "거래일",
                "value": "금액",
                "variable": "항목"
            },
            color_discrete_map={
                "영업매출": "skyblue",
                "비용": "red"
            }
        )
        st.plotly_chart(fig_daily, use_container_width=True)
    else:
        st.info("선택된 기간에 해당하는 일별 데이터가 없습니다.")

    # --------------------------------------------------
    # 6. Expense structure
    # --------------------------------------------------

    st.divider()
    st.header("6. 요식업 비용 구조 - 대분류")

    expense_start, expense_end = date_range_selector(
        "비용 구조",
        "expense_structure",
        data_min_date,
        data_max_date,
        data_min_date,
        data_max_date
    )

    expense_df = filter_by_date(base_df, expense_start, expense_end)

    main_expense = (
        expense_df[expense_df["비용"] > 0]
        .groupby("대분류")["비용"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    if not main_expense.empty:
        fig_main_expense = px.pie(
            main_expense,
            names="대분류",
            values="비용",
            labels={
                "대분류": "대분류",
                "비용": "비용"
            }
        )
        st.plotly_chart(fig_main_expense, use_container_width=True)
    else:
        st.info("선택된 기간에 해당하는 비용 데이터가 없습니다.")

    # --------------------------------------------------
    # Uncategorized transactions
    # --------------------------------------------------

    st.divider()
    st.header("기타 / 미분류 거래내역 확인")

    uncategorized_df = base_df[
        (base_df["대분류"] == "기타") |
        (base_df["세부분류"] == "미분류")
    ].copy()

    st.write(f"기타 / 미분류 거래 건수: {len(uncategorized_df):,}건")

    uncategorized_display_cols = [
        "거래일시",
        "적요",
        "기재내용",
        "거래구분",
        "입금액",
        "출금액",
        "거래후 잔액",
        "취급기관",
        "메모",
        "source_file"
    ]

    uncategorized_display_cols = [
        col for col in uncategorized_display_cols
        if col in uncategorized_df.columns
    ]

    if not uncategorized_df.empty:
        st.dataframe(
            uncategorized_df[uncategorized_display_cols]
            .sort_values("거래일시", ascending=False),
            use_container_width=True
        )

        uncategorized_csv = uncategorized_df.to_csv(
            index=False,
            encoding="utf-8-sig"
        )

        st.download_button(
            label="기타 / 미분류 거래내역 CSV 다운로드",
            data=uncategorized_csv,
            file_name="uncategorized_transactions.csv",
            mime="text/csv"
        )
    else:
        st.success("기타 / 미분류 거래내역이 없습니다.")

    # --------------------------------------------------
    # Minimal transaction summary
    # --------------------------------------------------

    st.divider()
    st.header("전체 거래내역 요약")

    s1, s2 = st.columns(2)

    s1.metric("총 거래 건수", f"{len(base_df):,}건")
    s2.metric("전체 거래 기간", f"{data_min_date} ~ {data_max_date}")

    cleaned_csv = base_df.to_csv(index=False, encoding="utf-8-sig")

    st.download_button(
        label="정리된 전체 데이터 CSV 다운로드",
        data=cleaned_csv,
        file_name="cleaned_woori_transactions.csv",
        mime="text/csv"
    )

except Exception as e:
    st.error("데이터를 불러오거나 처리하는 중 오류가 발생했습니다.")
    st.exception(e)