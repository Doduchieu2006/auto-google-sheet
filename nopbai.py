import time
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tagui import *

# ===============================
# 1. KẾT NỐI GOOGLE SHEET
# ===============================
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'service.json'

creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)
client = gspread.authorize(creds)

# ===============================
# 2. MỞ SHEET NopBai
# ===============================
SHEET_ID = '1YteY73LkEu2CbNEwWW8ghU5ozEAt28mRrkuLySDzAWk'
sheet = client.open_by_key(SHEET_ID)

ws = None
for w in sheet.worksheets():
    if w.title.strip().lower() == 'nopbai':
        ws = w
        break
if ws is None:
    raise Exception('❌ Không tìm thấy sheet NopBai')

# ===============================
# 3. ĐỌC DỮ LIỆU TỪ SHEET
# ===============================
data = ws.get_all_records()
df = pd.DataFrame(data)

df.columns = (
    df.columns
    .str.strip()
    .str.replace('\u00a0', '', regex=True)
)

if 'Nhóm' not in df.columns or 'Tình trạng' not in df.columns:
    raise Exception("❌ Sheet phải có cột 'Nhóm' và 'Tình trạng'")

# Lấy danh sách nhóm KHÔNG TRÙNG, GIỮ THỨ TỰ
groups_raw = [str(x).strip() for x in df['Nhóm'].dropna().tolist() if str(x).strip()]
groups = list(dict.fromkeys(groups_raw))

# ===============================
# 4. MỞ GOOGLE DRIVE (TAGUI)
# ===============================
init(visual_automation=True)
url('https://drive.google.com/drive/u/0/folders/1rZaTbI573-FgPzrbmEFrAq2lG0QgA8bs')
time.sleep(6)

# ===============================
# 5. KIỂM TRA NỘP BÀI (CHỈ CẦN CÓ TÊN NHÓM)
# ===============================
result = {}

# Ô search (thêm fallback tiếng Anh cho chắc)
SEARCH_BOX = '//*[@aria-label="Tìm trong Drive" or @aria-label="Search in Drive"]'

for group_name in groups:
    print(f'🔍 Kiểm tra nhóm: {group_name}')

    # ✅ Reset bằng ESC như bạn muốn
    keyboard('[esc]')
    time.sleep(0.5)

    # ✅ GÕ CHUẨN TAGUI: type(xpath, text)  -> hết lỗi text missing
    type(SEARCH_BOX, group_name)
    time.sleep(3)

    # ✅ Check “chỉ cần có tên nhóm” trong kết quả (case-insensitive)
    key_lower = group_name.lower()
    found_xpath = (
        "//div[@role='main']"
        "//*[@aria-label and "
        f"contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{key_lower}')]"
    )

    if exist(found_xpath):
        print(f'✅ {group_name}: Đã nộp')
        result[group_name] = 'Đã nộp'
    else:
        print(f'❌ {group_name}: Chưa nộp')
        result[group_name] = 'Chưa nộp'

# ===============================
# 6. GHI KẾT QUẢ VÀO SHEET
# ===============================
status_col = df.columns.get_loc('Tình trạng') + 1

for i, row in df.iterrows():
    group = str(row['Nhóm']).strip()
    ws.update_cell(i + 2, status_col, result.get(group, 'Chưa nộp'))

# ===============================
# 7. KẾT THÚC
# ===============================
close()
print('🎉 HOÀN TẤT KIỂM TRA NỘP BÀI')
