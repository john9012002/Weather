"""
Bước 3 — Dựng Dashboard Excel từ kết quả SQL (output/*.csv).

Sheet:
  Dashboard   — KPI + 4 biểu đồ (ngày/đêm theo tháng, chênh lệch, lượng mưa theo tháng, MA7)
  Thang       — tổng hợp tháng, toàn bộ bằng công thức AVERAGEIFS/SUMIFS/COUNTIFS
  Mua         — so sánh mùa mưa/khô + xác suất mưa (Markov 1 bậc) bằng COUNTIFS
  Data_Ngay   — dữ liệu ngày đã làm giàu từ SQL (MA7, LAG, spike, ngày/đêm)
  Cuc_doan    — danh sách ngày cực đoan (top 10%)
  Ghi_chu     — nguồn dữ liệu, định nghĩa

Chạy:  python build_dashboard.py            → HCM_Weather_Dashboard.xlsx
       python build_dashboard.py --demo     (thêm băng cảnh báo DỮ LIỆU GIẢ LẬP)
"""
import argparse
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
OUT = BASE / "output"

FONT = "Arial"
C_DARK, C_RAIN, C_HOT, C_COOL, C_GREY = "1F3A5F", "2E86C1", "D35400", "2874A6", "F2F4F7"
F_TITLE = Font(name=FONT, size=16, bold=True, color=C_DARK)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BODY = Font(name=FONT, size=10)
F_KPI = Font(name=FONT, size=18, bold=True, color=C_DARK)
F_KPI_LBL = Font(name=FONT, size=9, color="5D6D7E")
FILL_HDR = PatternFill("solid", fgColor=C_DARK)
FILL_KPI = PatternFill("solid", fgColor=C_GREY)
FILL_WARN = PatternFill("solid", fgColor="FDEDEC")
THIN = Side(style="thin", color="D5D8DC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header(ws, row, labels, widths=None):
    for j, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=j, value=lab)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if widths:
            ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    ws.row_dimensions[row].height = 32


def body_style(ws, r1, r2, c1, c2, fmt=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.font, cell.border = F_BODY, BORDER
            if fmt:
                cell.number_format = fmt


def warn_banner(ws, row, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1,
                value="⚠ DỮ LIỆU GIẢ LẬP để xem bố cục — chạy crawl_weather.py → run_analysis.py → "
                      "build_dashboard.py để có số liệu thật từ Open-Meteo.")
    c.font = Font(name=FONT, size=10, bold=True, color="C0392B")
    c.fill = FILL_WARN


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--out", default=str(BASE / "HCM_Weather_Dashboard.xlsx"))
    args = ap.parse_args()

    daily = pd.read_csv(OUT / "v_daily_enriched.csv")
    dn = pd.read_csv(OUT / "v_day_night.csv")
    ext = pd.read_csv(OUT / "v_extreme_days.csv")
    df = daily.merge(dn, on="date", how="left")
    df["date"] = pd.to_datetime(df["date"])
    df["year_month"] = df["date"].dt.strftime("%Y-%m")
    months = df["year_month"].drop_duplicates().tolist()

    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_m = wb.create_sheet("Thang")
    ws_s = wb.create_sheet("Mua")
    ws_d = wb.create_sheet("Data_Ngay")
    ws_x = wb.create_sheet("Cuc_doan")
    ws_n = wb.create_sheet("Ghi_chu")

    # ------------------------------------------------------------ Data_Ngay
    cols = [
        ("Ngày", "date", 12, "dd/mm/yyyy"), ("Năm-Tháng", "year_month", 10, None),
        ("Tháng", "month", 7, "0"), ("Mùa", "season", 10, None),
        ("Nhiệt độ max (°C)", "temp_max", 10, "0.0"), ("Nhiệt độ min (°C)", "temp_min", 10, "0.0"),
        ("Nhiệt độ TB (°C)", "temp_mean", 10, "0.0"), ("MA7 nhiệt độ TB (°C)", "temp_mean_ma7", 11, "0.00"),
        ("TB ban ngày 6h–18h (°C)", "temp_day_avg", 11, "0.0"),
        ("TB ban đêm 18h–6h (°C)", "temp_night_avg", 11, "0.0"),
        ("Chênh ngày–đêm (°C)", "day_night_gap", 10, "0.0"),
        ("Biên độ max–min (°C)", "diurnal_range", 10, "0.0"),
        ("Độ ẩm TB (%)", "rh_mean", 9, "0"), ("Lượng mưa (mm)", "precip_mm", 9, "0.0"),
        ("Ngày mưa (≥1mm)", "is_wet", 8, "0"), ("Mai mưa?", "is_wet_next_day", 8, "0"),
        ("Spike nhiệt", "is_spike", 8, "0"), ("Δ so hôm qua (°C)", "delta_vs_yesterday", 10, "+0.0;-0.0;0.0"),
        ("Δ so TB 7 ngày trước (°C)", "delta_vs_prev7", 11, "+0.0;-0.0;0.0"),
    ]
    header(ws_d, 1, [c[0] for c in cols], [c[2] for c in cols])
    for i, row in enumerate(df[[c[1] for c in cols]].itertuples(index=False), start=2):
        for j, v in enumerate(row, 1):
            ws_d.cell(row=i, column=j, value=None if pd.isna(v) else
                      (v.to_pydatetime() if hasattr(v, "to_pydatetime") else v))
    n_last = len(df) + 1
    for j, c in enumerate(cols, 1):
        body_style(ws_d, 2, n_last, j, j, c[3])
    ws_d.freeze_panes = "B2"
    ws_d.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{n_last}"
    ws_d.conditional_formatting.add(
        f"Q2:Q{n_last}", CellIsRule(operator="equal", formula=["1"],
                                    fill=PatternFill("solid", fgColor="FADBD8")))

    D = lambda col: f"Data_Ngay!${col}$2:${col}${n_last}"   # noqa: E731

    # ------------------------------------------------------------ Thang
    r0 = 3 if args.demo else 2
    if args.demo:
        warn_banner(ws_m, 1, 14)
    m_labels = ["Năm-Tháng", "Tháng", "Mùa", "Số ngày", "Max TB (°C)", "Min TB (°C)",
                "Ban ngày TB (°C)", "Ban đêm TB (°C)", "Chênh ngày–đêm (°C)", "Độ ẩm TB (%)",
                "Tổng mưa (mm)", "Số ngày mưa", "Khả năng mưa", "Số ngày spike"]
    header(ws_m, r0, m_labels, [11, 7, 10, 8, 10, 10, 11, 11, 11, 10, 11, 9, 10, 9])
    for k, ym in enumerate(months):
        r = r0 + 1 + k
        ws_m.cell(row=r, column=1, value=ym)
        ws_m[f"B{r}"] = f"=INDEX({D('C')},MATCH($A{r},{D('B')},0))"
        ws_m[f"C{r}"] = f"=INDEX({D('D')},MATCH($A{r},{D('B')},0))"
        ws_m[f"D{r}"] = f"=COUNTIFS({D('B')},$A{r})"
        ws_m[f"E{r}"] = f"=AVERAGEIFS({D('E')},{D('B')},$A{r})"
        ws_m[f"F{r}"] = f"=AVERAGEIFS({D('F')},{D('B')},$A{r})"
        ws_m[f"G{r}"] = f"=AVERAGEIFS({D('I')},{D('B')},$A{r})"
        ws_m[f"H{r}"] = f"=AVERAGEIFS({D('J')},{D('B')},$A{r})"
        ws_m[f"I{r}"] = f"=G{r}-H{r}"
        ws_m[f"J{r}"] = f"=AVERAGEIFS({D('M')},{D('B')},$A{r})"
        ws_m[f"K{r}"] = f"=SUMIFS({D('N')},{D('B')},$A{r})"
        ws_m[f"L{r}"] = f"=COUNTIFS({D('B')},$A{r},{D('O')},1)"
        ws_m[f"M{r}"] = f"=IF(D{r}=0,0,L{r}/D{r})"
        ws_m[f"N{r}"] = f"=COUNTIFS({D('B')},$A{r},{D('Q')},1)"
    m1, m2 = r0 + 1, r0 + len(months)
    for j, fmt in enumerate([None, "0", None, "0", "0.0", "0.0", "0.0", "0.0", "0.0",
                             "0", "#,##0.0", "0", "0%", "0"], 1):
        body_style(ws_m, m1, m2, j, j, fmt)
    ws_m.freeze_panes = ws_m.cell(row=m1, column=2)

    # ------------------------------------------------------------ Mua
    if args.demo:
        warn_banner(ws_s, 1, 9)
    header(ws_s, r0, ["Mùa", "Số ngày", "Khả năng mưa P(mưa)", "P(mai mưa | nay mưa)",
                      "P(mai mưa | nay khô)", "Max TB (°C)", "Độ ẩm TB (%)", "Tổng mưa (mm)",
                      "Chênh ngày–đêm TB (°C)"], [14, 9, 13, 14, 14, 10, 10, 12, 13])
    for k, s in enumerate(["Mùa mưa", "Mùa khô"]):
        r = r0 + 1 + k
        ws_s.cell(row=r, column=1, value=s)
        ws_s[f"B{r}"] = f"=COUNTIFS({D('D')},$A{r})"
        ws_s[f"C{r}"] = f"=COUNTIFS({D('D')},$A{r},{D('O')},1)/B{r}"
        ws_s[f"D{r}"] = (f"=IFERROR(COUNTIFS({D('D')},$A{r},{D('O')},1,{D('P')},1)/"
                         f"COUNTIFS({D('D')},$A{r},{D('O')},1,{D('P')},\">=0\"),0)")
        ws_s[f"E{r}"] = (f"=IFERROR(COUNTIFS({D('D')},$A{r},{D('O')},0,{D('P')},1)/"
                         f"COUNTIFS({D('D')},$A{r},{D('O')},0,{D('P')},\">=0\"),0)")
        ws_s[f"F{r}"] = f"=AVERAGEIFS({D('E')},{D('D')},$A{r})"
        ws_s[f"G{r}"] = f"=AVERAGEIFS({D('M')},{D('D')},$A{r})"
        ws_s[f"H{r}"] = f"=SUMIFS({D('N')},{D('D')},$A{r})"
        ws_s[f"I{r}"] = f"=AVERAGEIFS({D('K')},{D('D')},$A{r})"
    rt = r0 + 3
    ws_s.cell(row=rt, column=1, value="Cả năm")
    ws_s[f"B{rt}"] = f"=COUNT({D('E')})"
    ws_s[f"C{rt}"] = f"=SUM({D('O')})/B{rt}"
    ws_s[f"D{rt}"] = f"=IFERROR(COUNTIFS({D('O')},1,{D('P')},1)/COUNTIFS({D('O')},1,{D('P')},\">=0\"),0)"
    ws_s[f"E{rt}"] = f"=IFERROR(COUNTIFS({D('O')},0,{D('P')},1)/COUNTIFS({D('O')},0,{D('P')},\">=0\"),0)"
    ws_s[f"F{rt}"] = f"=AVERAGE({D('E')})"
    ws_s[f"G{rt}"] = f"=AVERAGE({D('M')})"
    ws_s[f"H{rt}"] = f"=SUM({D('N')})"
    ws_s[f"I{rt}"] = f"=AVERAGE({D('K')})"
    for j, fmt in enumerate([None, "0", "0.0%", "0.0%", "0.0%", "0.0", "0", "#,##0.0", "0.00"], 1):
        body_style(ws_s, r0 + 1, rt, j, j, fmt)
    for c in range(1, 10):
        ws_s.cell(row=rt, column=c).font = Font(name=FONT, size=10, bold=True)
    ws_s.cell(row=rt + 2, column=1,
              value="Cách đọc: P(mai mưa | nay mưa) là xác suất ngày hôm sau có mưa khi hôm nay đã mưa "
                    "(chuỗi Markov 1 bậc) — một mô hình dự báo khí hậu đơn giản, không thay thế dự báo thời tiết."
              ).font = Font(name=FONT, size=9, italic=True, color="5D6D7E")

    # ------------------------------------------------------------ Cuc_doan
    x_cols = [("Ngày", "date", 12), ("Mùa", "season", 10), ("Max (°C)", "temp_max", 9),
              ("Min (°C)", "temp_min", 9), ("Độ ẩm TB (%)", "rh_mean", 10),
              ("Độ ẩm min (%)", "rh_min", 10), ("Mưa (mm)", "precip_mm", 9),
              ("Nhãn cực đoan", "extreme_tags", 34)]
    header(ws_x, 1, [c[0] for c in x_cols], [c[2] for c in x_cols])
    ext = ext.sort_values("date")
    for i, row in enumerate(ext[[c[1] for c in x_cols]].itertuples(index=False), start=2):
        for j, v in enumerate(row, 1):
            ws_x.cell(row=i, column=j, value=None if pd.isna(v) else v)
    body_style(ws_x, 2, len(ext) + 1, 1, len(x_cols))
    ws_x.freeze_panes = "A2"

    # ------------------------------------------------------------ Dashboard
    ws = ws_dash
    ws.sheet_view.showGridLines = False
    for col in range(1, 21):
        ws.column_dimensions[get_column_letter(col)].width = 10.5
    ws["A1"] = "Khí hậu TP.HCM — biến động nhiệt độ & khả năng mưa theo mùa"
    ws["A1"].font = F_TITLE
    ws["A2"] = (f'="Dữ liệu ngày "&TEXT(MIN({D("A")}),"dd/mm/yyyy")&" → "&'
                f'TEXT(MAX({D("A")}),"dd/mm/yyyy")&" · Nguồn: Open-Meteo Archive API (ERA5)"')
    ws["A2"].font = F_KPI_LBL
    if args.demo:
        warn_banner(ws, 3, 20)

    kpis = [
        ("Nóng nhất (°C)", f"=MAX({D('E')})",
         f'=TEXT(INDEX({D("A")},MATCH(MAX({D("E")}),{D("E")},0)),"dd/mm/yyyy")', "0.0"),
        ("Mát nhất (°C)", f"=MIN({D('F')})",
         f'=TEXT(INDEX({D("A")},MATCH(MIN({D("F")}),{D("F")},0)),"dd/mm/yyyy")', "0.0"),
        ("Tổng lượng mưa (mm)", f"=SUM({D('N')})", f'=COUNTIFS({D("O")},1)&" ngày mưa"', "#,##0"),
        ("Khả năng mưa mùa mưa", f"=Mua!C{r0+1}", f'="Mùa khô: "&TEXT(Mua!C{r0+2},"0%")', "0%"),
        ("Chênh ngày–đêm TB (°C)", f"=AVERAGE({D('K')})",
         f'="Lớn nhất: "&FIXED(MAX({D("K")}),1)&" °C"', "0.0"),
        ("Số ngày spike nhiệt", f"=COUNTIFS({D('Q')},1)", '="Δ ≥ 2 °C"', "0"),
    ]
    kr = 5
    for k, (lab, f_val, f_sub, fmt) in enumerate(kpis):
        c0 = 1 + k * 3
        ws.merge_cells(start_row=kr, start_column=c0, end_row=kr, end_column=c0 + 2)
        ws.merge_cells(start_row=kr + 1, start_column=c0, end_row=kr + 1, end_column=c0 + 2)
        ws.merge_cells(start_row=kr + 2, start_column=c0, end_row=kr + 2, end_column=c0 + 2)
        a = ws.cell(row=kr, column=c0, value=lab)
        b = ws.cell(row=kr + 1, column=c0, value=f_val)
        s = ws.cell(row=kr + 2, column=c0, value=f_sub)
        a.font, b.font, s.font = F_KPI_LBL, F_KPI, F_KPI_LBL
        b.number_format = fmt
        for rr in range(kr, kr + 3):
            for cc in range(c0, c0 + 3):
                ws.cell(row=rr, column=cc).fill = FILL_KPI
        for x in (a, b, s):
            x.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[kr + 1].height = 28

    cats = Reference(ws_m, min_col=1, min_row=m1, max_row=m2)

    # Chart 1 — Nhiệt độ ban ngày vs ban đêm theo tháng
    ch1 = LineChart()
    ch1.title = "Nhiệt độ TB ban ngày vs ban đêm theo tháng"
    ch1.y_axis.title, ch1.y_axis.number_format = "°C", "0"
    ch1.add_data(Reference(ws_m, min_col=7, max_col=8, min_row=r0, max_row=m2), titles_from_data=True)
    ch1.set_categories(cats)
    for ser, color in zip(ch1.series, [C_HOT, C_COOL]):
        ser.graphicalProperties.line.solidFill = color
        ser.graphicalProperties.line.width = 28000
        ser.smooth = False
    ch1.y_axis.scaling.min, ch1.y_axis.scaling.max = 20, 35
    ch1.height, ch1.width = 7.5, 15
    ch1.legend.position = "b"
    ws.add_chart(ch1, "A9")

    # Chart 2 — Chênh lệch ngày–đêm (cột)
    ch2 = BarChart()
    ch2.title = "Chênh lệch nhiệt độ ngày – đêm theo tháng"
    ch2.y_axis.title, ch2.y_axis.number_format = "°C", "0.0"
    ch2.add_data(Reference(ws_m, min_col=9, min_row=r0, max_row=m2), titles_from_data=True)
    ch2.set_categories(cats)
    ch2.series[0].graphicalProperties.solidFill = C_HOT
    ch2.legend = None
    ch2.height, ch2.width = 7.5, 15
    ws.add_chart(ch2, "J9")

    # Chart 3 — Phân bổ lượng mưa theo tháng
    ch3 = BarChart()
    ch3.title = "Tổng lượng mưa theo tháng"
    ch3.y_axis.title, ch3.y_axis.number_format = "mm", "#,##0"
    ch3.add_data(Reference(ws_m, min_col=11, min_row=r0, max_row=m2), titles_from_data=True)
    ch3.set_categories(cats)
    ch3.series[0].graphicalProperties.solidFill = C_RAIN
    ch3.legend = None
    ch3.height, ch3.width = 7.5, 15
    ws.add_chart(ch3, "A25")

    # Chart 4 — Khả năng mưa (số ngày mưa / số ngày) theo tháng
    ch4 = BarChart()
    ch4.title = "Khả năng mưa theo tháng (tỉ lệ ngày ≥ 1 mm)"
    ch4.y_axis.number_format, ch4.y_axis.scaling.min, ch4.y_axis.scaling.max = "0%", 0, 1
    ch4.add_data(Reference(ws_m, min_col=13, min_row=r0, max_row=m2), titles_from_data=True)
    ch4.set_categories(cats)
    ch4.series[0].graphicalProperties.solidFill = C_COOL
    ch4.legend = None
    ch4.height, ch4.width = 7.5, 15
    ws.add_chart(ch4, "J25")

    # Chart 5 — Nhiệt độ TB ngày + MA7
    ch5 = LineChart()
    ch5.title = "Nhiệt độ trung bình ngày & trung bình động 7 ngày"
    ch5.y_axis.title, ch5.y_axis.number_format = "°C", "0"
    ch5.add_data(Reference(ws_d, min_col=7, max_col=8, min_row=1, max_row=n_last), titles_from_data=True)
    ch5.set_categories(Reference(ws_d, min_col=1, min_row=2, max_row=n_last))
    # Trục danh mục (không dùng DateAxis: openpyxl ghép sai axId → Excel "Repair" và xoá hết biểu đồ)
    ch5.x_axis.number_format = "mm/yy"
    ch5.x_axis.tickLblSkip, ch5.x_axis.tickMarkSkip = 30, 30
    ch5.series[0].graphicalProperties.line.solidFill = "AEB6BF"
    ch5.series[0].graphicalProperties.line.width = 9000
    ch5.series[1].graphicalProperties.line.solidFill = C_HOT
    ch5.series[1].graphicalProperties.line.width = 28000
    for ser in ch5.series:
        ser.smooth = False
    ch5.y_axis.scaling.min, ch5.y_axis.scaling.max = 22, 34
    ch5.height, ch5.width = 8, 30.5
    ch5.legend.position = "b"
    ws.add_chart(ch5, "A41")

    # ------------------------------------------------------------ Ghi_chu
    notes = [
        ("Nguồn", "Open-Meteo Historical Weather API (archive-api.open-meteo.com), tái phân tích ERA5; "
                  "toạ độ 10.8231, 106.6297 (TP.HCM); múi giờ UTC+7."),
        ("Mùa", "Mùa mưa: tháng 5–11; mùa khô: tháng 12–4."),
        ("Ngày mưa", "Lượng mưa ngày ≥ 1 mm (ngưỡng WMO)."),
        ("Ban ngày / đêm", "Ban ngày 06:00–17:59, ban đêm 18:00–05:59 trong cùng ngày dương lịch, tính từ dữ liệu giờ."),
        ("MA7", "AVG(temp_mean) OVER (ORDER BY date ROWS BETWEEN 6 PRECEDING AND CURRENT ROW); 6 ngày đầu có < 7 điểm."),
        ("Spike nhiệt", "temp_max − LAG(temp_max) ≥ 2 °C hoặc temp_max − TB 7 ngày trước ≥ 2 °C."),
        ("Cực đoan", "Top 10% theo PERCENT_RANK (nóng / mát / ẩm / khô)."),
        ("Khả năng mưa", "P(mưa) = số ngày mưa / số ngày. P(mai mưa | nay mưa) dùng LEAD(is_wet) — Markov 1 bậc."),
        ("Ô có thể sửa", "Chỉ sheet Data_Ngay là dữ liệu đầu vào; Thang, Mua, Dashboard tự tính lại bằng công thức."),
    ]
    ws_n.column_dimensions["A"].width, ws_n.column_dimensions["B"].width = 16, 110
    header(ws_n, 1, ["Mục", "Mô tả"])
    for i, (a, b) in enumerate(notes, start=2):
        ws_n.cell(row=i, column=1, value=a)
        ws_n.cell(row=i, column=2, value=b).alignment = Alignment(wrap_text=True, vertical="top")
    body_style(ws_n, 2, len(notes) + 1, 1, 2)

    for ch in (ch1, ch2, ch3, ch4, ch5):   # Excel 365 ẩn trục nếu delete mặc định
        ch.x_axis.delete = False
        ch.y_axis.delete = False

    wb.save(args.out)
    print(f"Đã tạo {args.out}")


if __name__ == "__main__":
    main()