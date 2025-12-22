
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ===============================
# 1. CẤU HÌNH
# ===============================
SERVICE_ACCOUNT_FILE = 'service.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# 🔹 SHEET NGUỒN (ĐIỂM DANH)
SOURCE_SPREADSHEET_ID = '1JjMO_2IBUEn0Yo6KmPWegGRdoOCd_-asFYAxNCrswgQ'
SOURCE_SHEET_NAME = 'Trang tính1'

# 🔹 SHEET ĐÍCH
TARGET_SPREADSHEET_ID = '1YteY73LkEu2CbNEwWW8ghU5ozEAt28mRrkuLySDzAWk'
TARGET_SHEET_NAME = 'Diemdanh'

# ===============================
# 2. KẾT NỐI GOOGLE SHEET
# ===============================
creds = ServiceAccountCredentials.from_json_keyfile_name(
    SERVICE_ACCOUNT_FILE, SCOPES
)
client = gspread.authorize(creds)

ws_src = client.open_by_key(SOURCE_SPREADSHEET_ID).worksheet(SOURCE_SHEET_NAME)
ws_dst = client.open_by_key(TARGET_SPREADSHEET_ID).worksheet(TARGET_SHEET_NAME)

# ===============================
# 3. ĐỌC DỮ LIỆU (KHÔNG DÙNG HEADER)
# ===============================
src_values = ws_src.get_all_values()[1:]   # bỏ header
dst_values = ws_dst.get_all_values()[1:]

# ===============================
# 4. MAP MSV → (ĐIỂM DANH, NGÀY VẮNG)
# ===============================
attendance_map = {}

for row in src_values:
    # cần tới cột V (index 21)
    if len(row) <= 21:
        continue

    msv = str(row[1]).strip().replace('.0', '')   # 🔥 MSV ở cột B (nguồn)
    src_ngay_vang = row[20].strip()               # cột U
    src_diem_danh = row[21].strip()               # cột V

    if not msv:
        continue

    attendance_map[msv] = {
        'Điểm danh': src_diem_danh,
        'Ngày vắng': src_ngay_vang
    }

# ===============================
# 5. GHI SANG SHEET ĐÍCH
#    A = MSV | D = Điểm danh | E = Ngày vắng
# ===============================
updates = []

for i, row in enumerate(dst_values):
    if len(row) < 1:
        continue

    msv = str(row[0]).strip().replace('.0', '')   # 🔥 MSV ở cột A (đích)

    if msv in attendance_map:
        updates.append({
            'range': f'D{i+2}',
            'values': [[attendance_map[msv]['Điểm danh']]]
        })
        updates.append({
            'range': f'E{i+2}',
            'values': [[attendance_map[msv]['Ngày vắng']]]
        })

# ===============================
# 6. BATCH UPDATE (ÉP GOOGLE SHEET HIỂU LÀ SỐ)
# ===============================
if updates:
    ws_dst.batch_update(
        updates,
        value_input_option='USER_ENTERED'
    )
    print(f'✅ ĐÃ UPDATE {len(updates)//2} SINH VIÊN')
else:
    print('⚠️ KHÔNG CÓ DỮ LIỆU ĐỂ UPDATE')
