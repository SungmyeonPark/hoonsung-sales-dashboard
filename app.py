import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime, timedelta
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


st.set_page_config(
    page_title="우리은행 가게 매출 분석 대시보드",
    layout="wide"
)

st.title("우리은행 거래내역 기반 요식업 매출 분석 대시보드")


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
# File helpers
# --------------------------------------------------

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


def list_drive_transaction_files(folder_id):
    """
    List xls, xlsx, csv files from a specific Google Drive folder.
    """

    service = get_drive_service()

    query = f"'{folder_id}' in parents and trashed = false"

    files = []
    page_token = None

    while True:
        results = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
            pageToken=page_token,
            pageSize=1000
        ).execute()

        files.extend(results.get("files", []))

        page_token = results.get("nextPageToken")

        if not page_token:
            break

    allowed_files = [
        file_info for file_info in files
        if file_info["name"].lower().endswith((".xls", ".xlsx", ".csv"))
    ]

    # 파일명 기준 정렬: 250508_260508.xls 같은 이름이면 자연스럽게 순서 정렬됨
    allowed_files = sorted(allowed_files, key=lambda x: x["name"])

    return allowed_files


def download_drive_file(file_id, file_name):
    """
    Download a Google Drive file into a BytesIO object.
    This object behaves like an uploaded file and can be used by pandas.
    """

    service = get_drive_service()

    request = service.files().get_media(fileId=file_id)
    file_buffer = BytesIO()

    downloader = MediaIoBaseDownload(file_buffer, request)

    done = False

    while not done:
        status, done = downloader.next_chunk()

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

def date_range_selector(label, key, min_date, max_date, default_start=None, default_end=None):
    """
    Reusable date range selector.
    User can choose either direct input or slider.
    """

    if default_start is None:
        default_start = min_date

    if default_end is None:
        default_end = max_date

    st.markdown(f"**{label} 기간 선택**")

    mode = st.radio(
        f"{label} 기간 선택 방식",
        ["직접 입력", "슬라이더"],
        horizontal=True,
        key=f"{key}_mode"
    )

    if mode == "직접 입력":
        c1, c2 = st.columns(2)

        with c1:
            start_date = st.date_input(
                "시작일",
                value=default_start,
                min_value=min_date,
                max_value=max_date,
                key=f"{key}_start"
            )

        with c2:
            end_date = st.date_input(
                "종료일",
                value=default_end,
                min_value=min_date,
                max_value=max_date,
                key=f"{key}_end"
            )

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

uploaded_files = load_files_from_google_drive()

if uploaded_files:
    try:
        all_dfs = []

        for uploaded_file in uploaded_files:
            temp_df = load_transaction_file(uploaded_file)
            temp_df = normalize_columns(temp_df)

            start_date, end_date = parse_woori_filename(uploaded_file.name)

            temp_df["source_file"] = uploaded_file.name
            temp_df["file_start_date"] = start_date
            temp_df["file_end_date"] = end_date

            all_dfs.append(temp_df)

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
            st.error(f"필수 컬럼이 없습니다: {missing_cols}")
            st.write("현재 인식된 컬럼명:")
            st.write(list(df.columns))
            st.stop()

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

        before_count = len(df)
        df = df.drop_duplicates(subset=duplicate_check_cols, keep="first")
        removed_count = before_count - len(df)

        if removed_count > 0:
            st.info(f"중복으로 보이는 거래 {removed_count:,}건을 제거했습니다.")

        # --------------------------------------------------
        # Sidebar filters
        # --------------------------------------------------

        st.sidebar.header("전체 필터")

        available_types = sorted(df["거래구분"].unique())
        selected_types = st.sidebar.multiselect(
            "거래구분 선택",
            available_types,
            default=available_types
        )

        available_main_categories = sorted(df["대분류"].unique())
        selected_main_categories = st.sidebar.multiselect(
            "대분류 선택",
            available_main_categories,
            default=available_main_categories
        )

        available_sub_categories = sorted(df["세부분류"].unique())
        selected_sub_categories = st.sidebar.multiselect(
            "세부분류 선택",
            available_sub_categories,
            default=available_sub_categories
        )

        available_files = sorted(df["source_file"].unique())
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

        summary_start, summary_end = date_range_selector(
            "주요 요약",
            "summary",
            data_min_date,
            data_max_date,
            data_min_date,
            data_max_date
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

        labor_cost = summary_df[
            summary_df["대분류"] == "인건비"
        ]["비용"].sum()

        rent_cost = summary_df[
            summary_df["대분류"] == "임대료"
        ]["비용"].sum()

        utility_cost = summary_df[
            summary_df["대분류"].isin(["공과금", "위생/관리비"])
        ]["비용"].sum()

        seafood_meat_ratio = seafood_meat_cost / operating_sales if operating_sales > 0 else 0
        general_food_ratio = general_food_cost / operating_sales if operating_sales > 0 else 0
        alcohol_ratio = alcohol_cost / operating_sales if operating_sales > 0 else 0
        labor_cost_ratio = labor_cost / operating_sales if operating_sales > 0 else 0
        rent_ratio = rent_cost / operating_sales if operating_sales > 0 else 0
        utility_ratio = utility_cost / operating_sales if operating_sales > 0 else 0

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

        total_food_alcohol_cost = seafood_meat_cost + general_food_cost + alcohol_cost
        total_food_alcohol_ratio = (
            total_food_alcohol_cost / operating_sales
            if operating_sales > 0
            else 0
        )

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

        period_start, period_end = date_range_selector(
            "월별/분기별 차트",
            "period_chart",
            data_min_date,
            data_max_date,
            data_min_date,
            data_max_date
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

        daily_start, daily_end = date_range_selector(
            "일별 추이",
            "daily_trend",
            daily_min_date,
            data_max_date,
            daily_min_date,
            data_max_date
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
        st.error("파일을 처리하는 중 오류가 발생했습니다.")
        st.exception(e)

else:
        st.info("Google Drive 폴더에 우리은행 거래내역 .xls, .xlsx, 또는 .csv 파일을 추가하세요.")