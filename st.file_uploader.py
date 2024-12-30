import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import gspread
import unicodedata
import traceback

platform_mapping = {
    '멜론': ['카카오', 'MelOn', '멜론'],
    '지니': ['KT뮤직NewGenie', '지니', '지니뮤직', 'Genie'],
    '벅스': ['벅스', 'Bugs', '벅스뮤직', 'BugsMusic', '벅스 뮤직', 'Bugs Music'],
    '소리바다': ['소리바다', 'SORIBADA', 'SORI BADA'],
    '유튜브뮤직': ['유튜브뮤직', '유튜브 뮤직', 'Youtube', '유튜브', '유튜브레드', 'Youtube Music', 'YoutubeMusic', 'YOUTUBE AD PARTNER', 'YOUTUBE RED'],
    '애플뮤직': ['Apple', '애플뮤직', '애플', 'Apple Music', 'AppleMusic', 'APPLE MUSIC'],
    '바이브': ['바이브', 'VIBE', '네이버', '네이버뮤직', 'Naver', 'Naver Music', 'NaverMusic'],
    '플로': ['플로', 'FLO'],
    '인스타그램': ['인스타그램', 'Instagram'],
    '페이스북': ['페이스북', 'Facebook'],
    '아마존': ['Amazon', '아마존', 'Amazon Prime (USD)'],
    'Resso': ['Resso', '레쏘'],
    'Deezer': ['Deezer', '디저'],
    'Tidal': ['Tidal', '타이달'],
    '컬러링': ['V컬러링'],
    '틱톡': ['틱톡', 'Tiktok', 'TikTok'],
    '카톡': ['카톡', '카카오톡'],
    '웨이브': ['WAVVE', '웨이버'],
    '올레뮤직': ['KT뮤직유선ollehMusic', '올레뮤직', 'ollehMusic'],
    '스포티파이': ['Spotify', '스포티파이']
}

artist_mapping = {
    '사운드힐즈': ['사운드힐즈', '사운드 힐즈', 'Soundhills', 'soundhills', 'Sound Hills', 'sound hills'],
    '스원': ['스원', 'Swon', 'swon', '스원(Swon)', '스원 (Swon)'],
    '나노말': ['나노말', 'NANOMAL', 'nanomal', 'Nanomal'],
    '이유카': ['이유카', 'Lee Yuka', 'LEE YUKA', 'Lee Yuka', 'LeeYuKa'],
    '하예지': ['하예지', '하예지 (발라드)'],
    '유마': ['유마'],
    '동자동휘': ['동자동휘'],
    '이규소': ['이규소'],
    '임광균': ['임광균'],
    '안우': ['안우', '안우 (Ahnoo)', '안우(Ahnoo)'],
    'Aaron': ['Aaron (댄스)', 'Aaron(댄스)', 'Aaron', 'aaron', 'AARON'],
    '위시스': ['위시스', 'Wiishes', 'WIISHES', 'wiishes']
}

def extract_spreadsheet_id(url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("유효하지 않은 Google Sheets URL입니다.")
    return match.group(1)

def get_google_credentials():
    credentials_dict = st.secrets["gcp"]
    credentials = Credentials.from_service_account_info(
        credentials_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return credentials

def normalize_text(text):
    return unicodedata.normalize('NFKC', str(text))

def mapping(name, map_dict):
    for key, keywords in map_dict.items():
        if any(keyword in name for keyword in keywords):
            return key
    return name

def read_excel_file(uploaded_file, header=None, sheet_name=0):
    if uploaded_file.name.endswith('.xls'):
        return pd.read_csv(uploaded_file, delimiter='\t', encoding='cp949')
    else:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header)

def process_files(uploaded_files):
    new_columns = ['정산월', '판매월', '유통사', '아티스트명', '앨범명', '곡명', '플랫폼', '서비스구분', '판매횟수', '매출', '정산금']
    final_df = pd.DataFrame(columns=new_columns)
    all_settlement_months = []

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        st.write(f"처리 중: {file_name}")

        # Extract company name
        try:
            company_name = extract_company_name(file_name)
        except ValueError as e:
            st.error(f"{file_name}: {e}")
            continue

        # Extract settlement month
        pattern = re.compile(r'(\d{4}-\d{2})_(.+)')
        match = pattern.match(file_name)
        if match:
            settlement_month = match.group(1)
            all_settlement_months.append(settlement_month)
            settlement_date = datetime.strptime(settlement_month, "%Y-%m")
            st.write(f"정산월: {settlement_month}")
        else:
            st.error(f"{file_name}: 파일명 형식이 맞지 않습니다.")
            continue

        company_name = normalize_text(company_name)
        st.write(f"Normalized company name: '{company_name}'")
        
        month_before_settlement = 3
        company_name_for_mapping = company_name

        # Determine header and data start rows based on company_name
        # (Same logic as your original script)
        if company_name == '라인엠컴퍼니':
            cutoff_date = datetime(2024, 12, 1)
            if settlement_date >= cutoff_date:
                header_row = 12
                data_start_row = 12
            else:
                header_row = 10
                data_start_row = 10
                company_name_for_mapping = '라인엠컴퍼니2'
        elif company_name == "뮤직앤뉴":
            month_before_settlement = 1
            header_row = 0
            data_start_row = 0
        elif company_name == "비스킷사운드":
            month_before_settlement = 1
            worksheet_name = '음원 상세내역'
            header_row = 3
            data_start_row = 3
        elif company_name == "미러볼뮤직":
            month_before_settlement = 1
            header_row = 0
            data_start_row = 0
        else:
            st.error(f"지원하지 않는 유통사입니다. {company_name}")
            continue

        days_before_settlement = month_before_settlement * 30
        sale_date = settlement_date - timedelta(days=days_before_settlement)
        sale_month = sale_date.strftime("%Y-%m")
        st.write(f"판매월: {sale_month}")

        # Read Excel file
        if company_name == "비스킷사운드":
            df = read_excel_file(uploaded_file, header=header_row, sheet_name=worksheet_name)
        else:
            df = read_excel_file(uploaded_file, header=header_row)

        df = df.iloc[(data_start_row - header_row):].reset_index(drop=True)

        # Define column mappings
        column_mapping = {
            '라인엠컴퍼니': {
                '아티스트명': '아티스트',
                '앨범명': '앨범명',
                '곡명': '곡명',
                '플랫폼': '서비스사',
                '서비스구분': '서비스종류',
                '판매횟수': 'HIT-s',
                '매출': '매출',
                '정산금': '아티스트정산금'
            },
            '라인엠컴퍼니2': {
                '아티스트명': '아티스트명',
                '앨범명': '앨범명',
                '곡명': '곡명',
                '플랫폼': '정산처',
                '서비스구분': '서비스명',
                '판매횟수': '카운트',
                '매출': '정산',
                '정산금': '계약자정산'
            },
            '뮤직앤뉴': {
                '아티스트명': '아티스트',
                '앨범명': '앨범명',
                '곡명': '곡명',
                '플랫폼': '사이트',
                '서비스구분': 'POC서비스명',
                '판매횟수': '판매횟수',
                '매출': '합계금액',
                '정산금': '권리사정산금액'
            },
            '비스킷사운드': {
                '아티스트명': '아티스트명',
                '앨범명': '앨범명',
                '곡명': '트랙명',
                '플랫폼': '서비스사이트',
                '서비스구분': 'MEDIA',
                '판매횟수': ['스트리밍', '다운로드', '기타수량'],
                '매출': '저작인접권료',
                '정산금': '인세',
            },
            '미러볼뮤직': {
                '아티스트명': '아티스트',
                '앨범명': '앨범명',
                '곡명': '곡명',
                '플랫폼': '',
                '서비스구분': '',
                '판매횟수': '',
                '매출': '합계금액',
                '정산금': '정산금액'
            }
        }

        data = []
        for index, row in df.iterrows():
            new_row = {}

            # Populate new_row based on mappings
            # (Same logic as your original script)

            # Example:
            new_row['정산월'] = settlement_month
            new_row['판매월'] = sale_month
            new_row['유통사'] = company_name
            new_row['아티스트명'] = mapping(row[column_mapping[company_name_for_mapping]['아티스트명']], artist_mapping)
            new_row['앨범명'] = row[column_mapping[company_name_for_mapping]['앨범명']]
            new_row['곡명'] = row[column_mapping[company_name_for_mapping]['곡명']]

            try:
                new_row['플랫폼'] = mapping(row[column_mapping[company_name_for_mapping]['플랫폼']], platform_mapping)
            except:
                new_row['플랫폼'] = ''

            try:
                new_row['서비스구분'] = row[column_mapping[company_name_for_mapping]['서비스구분']]
            except:
                new_row['서비스구분'] = ''

            for col in ['판매횟수', '매출', '정산금']:
                if column_mapping[company_name_for_mapping].get(col):
                    mapping_value = column_mapping[company_name_for_mapping].get(col)
                    if isinstance(mapping_value, list):
                        new_row[col] = sum([row[val] for val in mapping_value if val in row and pd.notna(row[val])])
                    else:
                        new_row[col] = row[mapping_value] if mapping_value in row else None
                else:
                    new_row[col] = ''

            data.append(new_row)

        temp_df = pd.DataFrame(data, columns=new_columns)
        temp_df.dropna(axis=1, how='all', inplace=True)

        if not temp_df.empty:
            final_df = pd.concat([final_df, temp_df], ignore_index=True)

    return final_df

def extract_company_name(file_name):
    pattern = re.compile(r'\d{4}-\d{2}_(.+?)\s*\(\d+\)?\.')
    match = pattern.match(file_name)
    if match:
        company_name = match.group(1)
    else:
        pattern = re.compile(r'\d{4}-\d{2}_(.+?)\.')
        match = pattern.match(file_name)
        if match:
            company_name = match.group(1)
        else:
            raise ValueError("파일명 형식이 맞지 않습니다.")
    return company_name

def append_to_google_sheets(final_df, spreadsheet_link, sheet_name):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_link)
    credentials = get_google_credentials()
    client = gspread.authorize(credentials)

    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
    except gspread.exceptions.SpreadsheetNotFound:
        raise ValueError("스프레드시트를 찾을 수 없습니다. 접근 권한이 없거나 스프레드시트 ID가 잘못되었습니다.")

    try:
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError("시트 이름이 잘못되었습니다.")

    existing_data = sheet.get_all_values()
    if not existing_data:
        existing_df = pd.DataFrame(columns=final_df.columns)
    else:
        existing_df = pd.DataFrame(existing_data[1:], columns=existing_data[0])

    existing_df_filtered = existing_df.drop(columns=['유통사_회사정산', '회사_아티스트정산'], errors='ignore').applymap(normalize_text)
    final_df_filtered = final_df.applymap(normalize_text)

    new_rows_df = final_df_filtered.loc[~final_df_filtered.apply(tuple, axis=1).isin(existing_df_filtered.apply(tuple, axis=1))]

    new_rows_df['유통사_회사정산'] = False
    new_rows_df['회사_아티스트정산'] = False

    new_rows = new_rows_df.values.tolist()

    if new_rows:
        sheet.append_rows(new_rows)
        st.success(f"{len(new_rows)}개의 새로운 행이 추가되었습니다.")
    else:
        st.info("추가할 새로운 행이 없습니다.")

st.title("정산 데이터 처리 및 Google Sheets 업로드")

uploaded_files = st.file_uploader("엑셀 파일을 업로드하세요", type=["xls", "xlsx"], accept_multiple_files=True)

spreadsheet_link = st.text_input("Google Sheets URL을 입력하세요")
sheet_name = st.text_input("Sheet 이름을 입력하세요", value="음원_정산내역(전체데이터)")

if st.button("처리 및 업로드 시작"):
    if uploaded_files and spreadsheet_link and sheet_name:
        try:
            final_df = process_files(uploaded_files)
            append_to_google_sheets(final_df, spreadsheet_link, sheet_name)
        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.text(traceback.format_exc())
    else:
        st.warning("모든 필드를 입력해주세요.")