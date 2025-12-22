from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
from datetime import datetime, timedelta
import unicodedata
import re

# ===============================
# 1. CẤU HÌNH
# ===============================
SERVICE_ACCOUNT_FILE = 'service.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1YteY73LkEu2CbNEwWW8ghU5ozEAt28mRrkuLySDzAWk'
SHEET_NAME = 'Youtube'

API_KEY = open('API_keys.txt').read().strip()
CHANNEL_ID = 'UCc_RGAKIULbK6MRvAu47YKQ'
CLASS_NAME = '64HTTT3'

# ===============================
# 2. HÀM CHUẨN HÓA
# ===============================
def normalize_text(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    s = s.lower()
    s = re.sub(r'[-_/]+', ' ', s)
    s = re.sub(r'[^0-9a-z\s]+', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def extract_group_number(group_name: str):
    s = normalize_text(group_name)
    m = re.search(r'nhom\s*(\d+)', s)
    return int(m.group(1)) if m else None


# ===============================
# 3. HÀM LẤY LINK THEO LỚP + NHÓM (1–12)
# ===============================
def get_links_by_class_and_group(videos, class_name, group_from=1, group_to=12):
    """
    Bắt được:
    nhom1, nhom 1, nhom_1, nhóm1, NHOM1...
    """
    result = {g: [] for g in range(group_from, group_to + 1)}
    norm_class = normalize_text(class_name)

    for v in videos:
        title_norm = normalize_text(v['title'])

        # phải có tên lớp
        if norm_class not in title_norm:
            continue

        for g in range(group_from, group_to + 1):
            # 🔥 REGEX CHUẨN: nhom + optional space + số
            pattern = rf'nhom\s*{g}\b'
            if re.search(pattern, title_norm):
                result[g].append(v['link'])

    return result



# ===============================
# 4. KẾT NỐI GOOGLE SHEET
# ===============================
creds = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
gc = gspread.authorize(creds)
ws = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

df = pd.DataFrame(ws.get_all_records())
df.columns = df.columns.str.strip()

required_cols = ['Nhóm', 'Tiêu đề', 'Link', 'Tình trạng']
for c in required_cols:
    if c not in df.columns:
        raise Exception(f"❌ Sheet thiếu cột: {c}")

# ===============================
# 5. LẤY VIDEO 30 NGÀY GẦN NHẤT
# ===============================
youtube = build('youtube', 'v3', developerKey=API_KEY)
thirty_days_ago = (datetime.utcnow() - timedelta(days=30)).isoformat("T") + "Z"

video_ids = []
next_page = None

while True:
    res = youtube.search().list(
        channelId=CHANNEL_ID,
        part='id',
        order='date',
        publishedAfter=thirty_days_ago,
        type='video',
        maxResults=50,
        pageToken=next_page
    ).execute()

    for item in res['items']:
        video_ids.append(item['id']['videoId'])

    next_page = res.get('nextPageToken')
    if not next_page:
        break

video_ids = list(set(video_ids))

# ===============================
# 6. LẤY TITLE + LINK VIDEO
# ===============================
videos = []

for i in range(0, len(video_ids), 50):
    chunk = video_ids[i:i+50]
    res = youtube.videos().list(
        part='snippet',
        id=','.join(chunk)
    ).execute()

    for v in res['items']:
        videos.append({
            'title': v['snippet']['title'],
            'link': f"https://www.youtube.com/watch?v={v['id']}"
        })

# ===============================
# 7. KIỂM TRA NHÓM 1–12 + LỚP
# ===============================
links_by_group = get_links_by_class_and_group(
    videos,
    CLASS_NAME,
    group_from=1,
    group_to=12
)

# ===============================
# 8. GHI VÀO SHEET (4 CỘT)
# ===============================
for i, row in df.iterrows():
    group_name = str(row['Nhóm']).strip()
    group_num = extract_group_number(group_name)

    if group_num and group_num in links_by_group and links_by_group[group_num]:
        titles = []
        links = []

        for v in videos:
            if v['link'] in links_by_group[group_num]:
                titles.append(v['title'])
                links.append(v['link'])

        ws.update_cell(i + 2, df.columns.get_loc('Tiêu đề') + 1, '\n'.join(titles))
        ws.update_cell(i + 2, df.columns.get_loc('Link') + 1, '\n'.join(links))
        ws.update_cell(i + 2, df.columns.get_loc('Tình trạng') + 1, 'Đã nộp')

        print(f'✅ Nhóm {group_num}: Đã nộp ({len(links)} video)')
    else:
        ws.update_cell(i + 2, df.columns.get_loc('Tiêu đề') + 1, '')
        ws.update_cell(i + 2, df.columns.get_loc('Link') + 1, '')
        ws.update_cell(i + 2, df.columns.get_loc('Tình trạng') + 1, 'Chưa nộp')

        print(f'❌ Nhóm {group_num}: Chưa nộp')

print('🎉 HOÀN TẤT CHECK YOUTUBE THEO LỚP + NHÓM')
