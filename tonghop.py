import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import cellFormat, numberFormat, format_cell_range

# ===============================
# 1. CẤU HÌNH
# ===============================
SERVICE_ACCOUNT_FILE = 'service.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1YteY73LkEu2CbNEwWW8ghU5ozEAt28mRrkuLySDzAWk'

SHEET_TONGHOP = 'Tonghop'
SHEET_NOPBAI = 'Nopbai'
SHEET_DIEMDANH = 'Diemdanh'
SHEET_QUIZ = 'Quiz'

# ===============================
# 2. KẾT NỐI GOOGLE SHEET
# ===============================
creds = ServiceAccountCredentials.from_json_keyfile_name(
    SERVICE_ACCOUNT_FILE, SCOPES
)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)

# ===============================
# 3. HÀM TIỆN ÍCH
# ===============================
def read_df_records(sheet_name):
    ws = sh.worksheet(sheet_name)
    df = pd.DataFrame(ws.get_all_records())
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_df_values(sheet_name):
    ws = sh.worksheet(sheet_name)
    values = ws.get_all_values()
    df = pd.DataFrame(values[1:], columns=values[0])
    df.columns = [str(c).strip() for c in df.columns]
    return df


def norm_msv(df):
    if 'MSV' in df.columns:
        df['MSV'] = (
            df['MSV']
            .astype(str)
            .str.strip()
            .replace('.0', '', regex=False)
        )
    return df


def parse_quiz_score(val):
    if not val:
        return 0
    val = str(val).strip()
    if '/' in val:
        try:
            return float(val.split('/')[0])
        except:
            return 0
    try:
        return float(val)
    except:
        return 0


def to_number(val):
    """
    Ép '15 hoặc 14,5 -> NUMBER thật
    """
    if val is None or val == "":
        return ""
    return float(str(val).replace(',', '.'))

# ===============================
# 4. ĐỌC DỮ LIỆU
# ===============================
df_tonghop = norm_msv(read_df_records(SHEET_TONGHOP))
df_nopbai = norm_msv(read_df_records(SHEET_NOPBAI))
df_quiz = norm_msv(read_df_records(SHEET_QUIZ))

# 🔥 Diemdanh PHẢI đọc bằng get_all_values
df_diemdanh = norm_msv(read_df_values(SHEET_DIEMDANH))

# ===============================
# 5. MAP THEO MSV
# ===============================
map_nopbai = dict(zip(df_nopbai['MSV'], df_nopbai['Tình trạng']))
map_quiz = dict(zip(df_quiz['MSV'], df_quiz['Điểm số']))
map_diemdanh = dict(zip(df_diemdanh['MSV'], df_diemdanh['Điểm danh']))
map_ngayvang = dict(zip(df_diemdanh['MSV'], df_diemdanh['Ngày vắng']))

# ===============================
# 6. FORMAT CỘT H & I VỀ NUMBER (CHẠY 1 LẦN CŨNG OK)
# ===============================
ws_tonghop = sh.worksheet(SHEET_TONGHOP)

fmt_number = cellFormat(
    numberFormat=numberFormat(
        type='NUMBER',
        pattern='0.##'   # 👈 BẮT BUỘC
    )
)

format_cell_range(ws_tonghop, 'H2:H1000', fmt_number)
format_cell_range(ws_tonghop, 'I2:I1000', fmt_number)

# ===============================
# 7. TỔNG HỢP DỮ LIỆU
# ===============================
output = []

for _, row in df_tonghop.iterrows():
    msv = row['MSV']

    quiz = parse_quiz_score(map_quiz.get(msv, ""))
    nopbai = map_nopbai.get(msv, "Chưa nộp")

    diem_danh_raw = map_diemdanh.get(msv, "")
    ngay_vang_raw = map_ngayvang.get(msv, "")

    output.append([
        quiz,                         # F - Quiz
        nopbai,                       # G - Nộp bài
        to_number(diem_danh_raw),     # H - Điểm danh (NUMBER)
        to_number(ngay_vang_raw)      # I - Ngày vắng (NUMBER)
    ])

# ===============================
# 8. UPDATE GOOGLE SHEET
# ===============================
start_row = 2
end_row = start_row + len(output) - 1

ws_tonghop.update(
    f'F{start_row}:I{end_row}',
    output,
    value_input_option='USER_ENTERED'
)

print("✅ HOÀN TẤT – KHÔNG CÒN `'15` – CỘT H/I LÀ NUMBER THẬT")
