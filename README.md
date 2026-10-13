# Phân tích biến động nhiệt độ & khả năng mưa theo mùa tại TP.HCM

Pipeline 3 bước: **Python (crawl) → SQL (hàm cửa sổ) → Excel (dashboard)**.

```bash
pip install -r requirements.txt
python crawl_weather.py      # 1. Open-Meteo → data/*.csv + weather.db (365 ngày, không cần API key)
python run_analysis.py       # 2. chạy analysis.sql, xuất output/*.csv
python build_dashboard.py    # 3. → HCM_Weather_Dashboard.xlsx
```

## Các file
| File | Nội dung |
|---|---|
| `crawl_weather.py` | Gọi Open-Meteo Archive API (dữ liệu ngày + giờ): nhiệt độ max/min/TB, độ ẩm, lượng mưa, số giờ mưa, gió, giật. Lưu CSV + SQLite. |
| `analysis.sql` | 5 view: `v_daily_enriched` (MA7 bằng `AVG() OVER`, `LAG`/`LEAD`, cờ spike), `v_extreme_days` (`PERCENT_RANK` top 10%), `v_day_night`, `v_monthly`, `v_rain_forecast_season` (Markov 1 bậc). Kèm truy vấn gaps-and-islands tìm chuỗi ngày mưa dài nhất. |
| `run_analysis.py` | Chạy SQL, xuất kết quả, in tóm tắt ra màn hình. |
| `build_dashboard.py` | Excel 6 sheet: Dashboard (6 KPI + 5 biểu đồ), Thang, Mua, Data_Ngay, Cuc_doan, Ghi_chu. Sheet tổng hợp dùng công thức `AVERAGEIFS/SUMIFS/COUNTIFS` — sửa Data_Ngay là dashboard tự cập nhật. |

## Định nghĩa
- Mùa mưa: tháng 5–11, mùa khô: tháng 12–4.
- Ngày mưa: ≥ 1 mm (ngưỡng WMO).
- Ban ngày 06:00–17:59, ban đêm 18:00–05:59.
- Spike nhiệt: `temp_max` tăng ≥ 2 °C so với hôm qua (`LAG`) hoặc so với TB 7 ngày trước.
- "Dự báo khả năng mưa": P(mưa) theo mùa và P(mai mưa | nay mưa) bằng `LEAD(is_wet)` — là xác suất khí hậu, không phải dự báo thời tiết.

## Lưu ý
- Open-Meteo Archive trễ ~5 ngày nên dữ liệu kết thúc 5 ngày trước ngày chạy (`--lag` để đổi).
- Tháng đầu và tháng cuối thường chỉ có một phần số ngày — khi trình bày nên ghi chú hoặc lọc bỏ.
- SQL viết cho SQLite; các dòng khác biệt với PostgreSQL được đánh dấu `[PG]`.
