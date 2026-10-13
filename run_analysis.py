"""
Bước 2b — Chạy analysis.sql trên weather.db và xuất kết quả các view ra CSV
(output/*.csv) để đưa vào Excel.

Chạy:  python run_analysis.py
"""
import sqlite3
from pathlib import Path

import pandas as pd

BASE = Path(__file__).parent
OUT = BASE / "output"
VIEWS = ["v_daily_enriched", "v_day_night", "v_monthly",
         "v_extreme_days", "v_rain_forecast_season"]


def main():
    OUT.mkdir(exist_ok=True)
    with sqlite3.connect(BASE / "weather.db") as con:
        con.executescript((BASE / "analysis.sql").read_text(encoding="utf-8"))
        for v in VIEWS:
            df = pd.read_sql(f"SELECT * FROM {v}", con)
            df.to_csv(OUT / f"{v}.csv", index=False, encoding="utf-8-sig")
            print(f"  {v:<24} {len(df):>5} dòng → output/{v}.csv")

        print("\nKhả năng mưa theo mùa:")
        print(pd.read_sql("SELECT * FROM v_rain_forecast_season", con).to_string(index=False))
        print("\n5 ngày tăng nhiệt đột ngột nhất:")
        print(pd.read_sql("""
            SELECT date, temp_max_prev_day AS hom_qua, temp_max AS hom_nay,
                   delta_vs_yesterday, delta_vs_prev7
            FROM v_daily_enriched WHERE is_spike = 1
            ORDER BY MAX(COALESCE(delta_vs_yesterday,0), COALESCE(delta_vs_prev7,0)) DESC
            LIMIT 5""", con).to_string(index=False))


if __name__ == "__main__":
    main()