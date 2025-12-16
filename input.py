import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# ==================================================
# 1. KẾT NỐI GOOGLE SHEETS
# ==================================================
def connect_sheet(sheet_id):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "service.json",
        scopes=scopes
    )
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id)


# ==================================================
# 2. ĐỌC DỮ LIỆU (KHÔNG MẤT DẤU ,)
# ==================================================
def read_data(spreadsheet, sheet_name="Data"):
    ws = spreadsheet.worksheet(sheet_name)
    values = ws.get_all_values()
    headers = values[0]
    rows = values[1:]
    return pd.DataFrame(rows, columns=headers)


# ==================================================
# 3. LÀM SẠCH & CHUẨN HÓA
# ==================================================
def clean_data(df):
    df = df.copy()
    df.columns = df.columns.str.strip()

    # Chuẩn hóa tên cột
    df = df.rename(columns={
        "MSV": "Mã sinh viên",
        "Chuyên cần": "Điểm danh",
        "Điểm QT": "Quá trình",
        "Điểm giữa kỳ": "Giữa kỳ",
        "Điểm cuối kỳ": "Cuối kỳ"
    })

    # Xóa dòng không có mã SV
    df = df[df["Mã sinh viên"].astype(str).str.strip() != ""]

    # Xóa trùng sinh viên
    df = df.drop_duplicates(subset=["Mã sinh viên"], keep="first")

    # Xử lý cột Nhóm (merge cell → fill xuống)
    if "Nhóm" in df.columns:
        df["Nhóm"] = (
            df["Nhóm"]
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .ffill()
        )
        df["Nhóm"] = pd.to_numeric(df["Nhóm"], errors="coerce")

    # Chuẩn hóa cột điểm
    score_cols = ["Điểm danh", "Quá trình", "Giữa kỳ", "Cuối kỳ"]
    for col in score_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


# ==================================================
# 4. TÍNH ĐIỂM TỔNG KẾT (LÀM TRÒN 1 CHỮ SỐ)
# ==================================================
def calculate_total(df):
    df = df.copy()
    df["Tổng kết"] = (
        0.1 * df["Điểm danh"]
        + 0.3 * df["Quá trình"]
        + 0.2 * df["Giữa kỳ"]
        + 0.4 * df["Cuối kỳ"]
    ).round(1)
    return df


# ==================================================
# 5. XẾP HẠNG CÁ NHÂN – TOÀN LỚP
# ==================================================
def ranking_individual(df):
    rank = (
        df.sort_values("Tổng kết", ascending=False)
        .reset_index(drop=True)
    )

    # COMPETITION RANKING: 1 2 2 4 4 6
    rank["Xếp hạng"] = (
        rank["Tổng kết"]
        .rank(method="min", ascending=False)
        .astype(int)
    )

    return rank

# ==================================================
# 6. XẾP HẠNG NHÓM (DÙNG NHÓM CÓ SẴN)
# ==================================================
def ranking_group(df):
    group_df = (
        df.groupby("Nhóm")["Tổng kết"]
        .mean()
        .reset_index()
        .rename(columns={"Tổng kết": "Điểm TB nhóm"})
    )

    group_df["Xếp hạng nhóm"] = (
        group_df["Điểm TB nhóm"]
        .rank(method="min", ascending=False)
        .astype(int)
    )

    return group_df.sort_values("Xếp hạng nhóm")


# ==================================================
# 7. GHI DATAFRAME RA GOOGLE SHEETS
# ==================================================
def write_sheet(spreadsheet, sheet_name, df):
    try:
        ws = spreadsheet.worksheet(sheet_name)
        ws.clear()
    except:
        ws = spreadsheet.add_worksheet(
            title=sheet_name, rows=1000, cols=30
        )

    ws.update(
        [df.columns.tolist()] + df.values.tolist(),
        value_input_option="USER_ENTERED"
    )


# ==================================================
# 8. MAIN
# ==================================================
def main():
    SHEET_ID = "1p_o86o2HqGVnb1-sapAtu-UaHxFKd1JqZz36Yr28lGY"

    spreadsheet = connect_sheet(SHEET_ID)

    # 1. Đọc & xử lý dữ liệu
    df = read_data(spreadsheet, "Data")
    df = clean_data(df)
    df = calculate_total(df)

    # 2. Xếp hạng TOÀN LỚP
    rank_ind = ranking_individual(df)

    # 3. Ghi sheet NHÓM
    rank_group = ranking_group(rank_ind)
    write_sheet(spreadsheet, "Ranking_Group", rank_group)

    # 4. Ghi sheet CÁ NHÂN – SẮP XẾP THEO THỨ TỰ XẾP HẠNG
    final_output = (
        rank_ind
        .sort_values(["Xếp hạng", "Tổng kết"], ascending=[True, False])
        .reset_index(drop=True)
        .fillna("")
    )

    write_sheet(spreadsheet, "Ranking_Individual", final_output)

    print("✅ HOÀN THÀNH – ĐÃ SẮP XẾP THEO XẾP HẠNG TOÀN LỚP")


if __name__ == "__main__":
    main()
