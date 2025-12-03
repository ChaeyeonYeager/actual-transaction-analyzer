import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================
# 엑셀 스타일 적용 함수
# =====================================================

def apply_style_to_range(ws, start_row, end_row, start_col, end_col,
                         header_row=None, total_rows=None, total_cols=None):
    """ 특정 범위에 스타일 적용 """
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    total_fill = PatternFill(start_color="DEEAF6", end_color="DEEAF6", fill_type="solid")
    total_font = Font(bold=True)

    thin = Side(border_style="thin", color="BFBFBF")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # 헤더
            if header_row and r == header_row:
                cell.fill = header_fill
                cell.font = header_font

            # 합계 행
            if total_rows and r in total_rows:
                cell.fill = total_fill
                cell.font = total_font

            # 합계 열
            if total_cols and c in total_cols:
                cell.fill = total_fill
                cell.font = total_font


def autosize_columns(ws):
    """열 너비 자동 조정 """
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        # 최소 폭 15, 넉넉한 여유 +3
        adjusted_width = max(max_len + 3, 15)
        ws.column_dimensions[col_letter].width = adjusted_width



# =====================================================
# 월별 표 + 년도별 표(세로형) 2개를 한 시트에 저장
# =====================================================

def save_excel_with_two_tables(prefix, df_monthly, df_yearly):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{timestamp}_{prefix}.xlsx"

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_monthly.to_excel(writer, index=False, startrow=0)
        start_row = len(df_monthly) + 3  # 2줄 띄우기
        df_yearly.to_excel(writer, index=False, startrow=start_row)

    wb = load_workbook(filename)
    ws = wb.active

    m_rows = len(df_monthly)
    m_cols = len(df_monthly.columns)

    y_rows = len(df_yearly)
    y_cols = len(df_yearly.columns)
    y_start = m_rows + 3

    # 월별 스타일
    apply_style_to_range(
        ws,
        start_row=1,
        end_row=m_rows + 1,
        start_col=1,
        end_col=m_cols,
        header_row=1,
        total_rows=[m_rows + 1],
        total_cols=[m_cols]
    )

    # 년도별 스타일
    apply_style_to_range(
        ws,
        start_row=y_start + 1,
        end_row=y_start + y_rows + 1,
        start_col=1,
        end_col=y_cols,
        header_row=y_start + 1,
        total_rows=[],   # ❗ 년도별 합계 제거 요청 → 빈 리스트
        total_cols=[]
    )

    autosize_columns(ws)
    wb.save(filename)
    return filename


# =====================================================
# 분석 1: 강남구 월/년도 분석
# =====================================================

def 분석_강남구(df):
    df = df[df['시군구'].str.contains("강남구")]

    df["동"] = df["시군구"].str.split().str[2]
    df["년월"] = df["계약년월"].astype(str)
    df["년도"] = df["년월"].str[:4]

    # 월별
    monthly = df.groupby(["동", "년월"]).size().reset_index(name="건수")
    pivot_m = monthly.pivot(index="동", columns="년월", values="건수").fillna(0).astype(int)
    pivot_m["합계"] = pivot_m.sum(axis=1)
    pivot_m.loc["합계"] = pivot_m.sum()

    # 년도별 (세로형) 합계 제거
    yearly = df.groupby(["동", "년도"]).size().reset_index(name="건수")
    pivot_y = yearly.pivot(index="동", columns="년도", values="건수").fillna(0).astype(int)

    return save_excel_with_two_tables("강남구_분석",
                                      pivot_m.reset_index(),
                                      pivot_y.reset_index())


# =====================================================
# 분석 2: 금액대별 TOP6
# (표 1개만 생성)
# =====================================================

def save_single_table(prefix, df):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{timestamp}_{prefix}.xlsx"
    df.to_excel(filename, index=False)

    wb = load_workbook(filename)
    ws = wb.active

    max_row = ws.max_row
    max_col = ws.max_column

    apply_style_to_range(
        ws,
        start_row=1,
        end_row=max_row,
        start_col=1,
        end_col=max_col,
        header_row=1,
        total_rows=[max_row],
        total_cols=[max_col]
    )

    autosize_columns(ws)
    wb.save(filename)
    return filename


def 분석_금액대(df):
    target_gu = ["강남구", "성동구", "종로구", "중구", "용산구", "마포구"]

    df["구"] = df["시군구"].str.extract(r"서울특별시 (\S+구)")
    df["금액"] = df["거래금액(만원)"].astype(str).replace(",", "", regex=True).astype(float) / 10000

    def 구간(x):
        if x < 50: return "50억 미만"
        elif x < 100: return "50~100억 미만"
        elif x < 200: return "100~200억 미만"
        elif x < 400: return "200~400억 미만"
        elif x < 1000: return "400억 이상"
        else: return "1000억 이상"

    df["금액구간"] = df["금액"].apply(구간)
    df = df[df["구"].isin(target_gu)]

    grouped = df.groupby(["구", "금액구간"]).size().reset_index(name="건수")
    pivot = grouped.pivot(index="구", columns="금액구간", values="건수").fillna(0).astype(int)

    pivot["합계"] = pivot.sum(axis=1)
    pivot.loc["합계"] = pivot.sum()

    return save_single_table("금액대별_6개구_분석", pivot.reset_index())


# =====================================================
# 분석 3: 서울시 월/년도 분석
# =====================================================

def 분석_서울시(df):
    df["년월"] = df["계약년월"].astype(str)
    df["년도"] = df["년월"].str[:4]

    # 월별
    monthly = df.groupby("년월").size().reset_index(name="건수")
    pivot_m = monthly.pivot_table(index=None, columns="년월", values="건수").fillna(0).astype(int)
    pivot_m["합계"] = pivot_m.sum(axis=1)

    # 년도별 (세로형, 합계 행 없음)
    yearly = df.groupby("년도").size().reset_index(name="건수")

    return save_excel_with_two_tables("서울시_분석", pivot_m, yearly)


# =====================================================
# 헤더 자동 찾기
# =====================================================

def read_excel_with_auto_header(path):
    temp = pd.read_excel(path, header=None)
    for i in range(len(temp)):
        if "시군구" in temp.iloc[i].astype(str).values:
            return pd.read_excel(path, header=i)
    raise ValueError("헤더 찾기 실패")


# =====================================================
# GUI
# =====================================================

selected_files = []

def add_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    for f in files:
        if f not in selected_files:
            selected_files.append(f)
            listbox.insert(tk.END, f)

def remove_file():
    sel = listbox.curselection()
    if sel:
        idx = sel[0]
        listbox.delete(idx)
        selected_files.pop(idx)

def start_analysis():
    if not selected_files:
        messagebox.showerror("오류", "파일을 선택하세요.")
        return

    do1 = var1.get()
    do2 = var2.get()
    do3 = var3.get()

    if not (do1 or do2 or do3):
        messagebox.showerror("오류", "최소 1개의 분석을 선택하세요.")
        return

    combined = pd.DataFrame()
    try:
        for f in selected_files:
            combined = pd.concat([combined, read_excel_with_auto_header(f)], ignore_index=True)

        outputs = []

        if do1:
            outputs.append(분석_강남구(combined.copy()))
        if do2:
            outputs.append(분석_금액대(combined.copy()))
        if do3:
            outputs.append(분석_서울시(combined.copy()))

        messagebox.showinfo("완료", "\n".join(outputs))

    except Exception as e:
        messagebox.showerror("에러", str(e))


root = tk.Tk()
root.title("실거래가 자동 분석기")

tk.Label(root, text="엑셀 파일을 선택하세요").pack(pady=5)

listbox = tk.Listbox(root, width=80, height=8)
listbox.pack()

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Button(frame, text="파일 추가 (+)", command=add_files).grid(row=0, column=0, padx=5)
tk.Button(frame, text="파일 삭제 (-)", command=remove_file).grid(row=0, column=1, padx=5)

tk.Label(root, text="실행할 분석 선택:").pack(pady=10)

var1 = tk.IntVar()
var2 = tk.IntVar()
var3 = tk.IntVar()

tk.Checkbutton(root, text="① 강남구 월/년도 분석", variable=var1).pack(anchor="w", padx=30)
tk.Checkbutton(root, text="② 금액대별 6개구 분석", variable=var2).pack(anchor="w", padx=30)
tk.Checkbutton(root, text="③ 서울시 월/년도 분석", variable=var3).pack(anchor="w", padx=30)

tk.Button(root, text="분석 시작", bg="#2e8b57", fg="black",
          font=("맑은 고딕", 10, "bold"), width=20,
          command=start_analysis).pack(pady=20)

root.mainloop()
