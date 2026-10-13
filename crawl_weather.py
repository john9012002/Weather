"""
Bước 1 — Crawl dữ liệu thời tiết TP.HCM (365 ngày gần nhất) từ Open-Meteo Archive API.

Đầu ra:
  data/weather_daily.csv   — 1 dòng / ngày
  data/weather_hourly.csv  — 1 dòng / giờ (dùng để tính nhiệt độ ban ngày vs ban đêm)
  weather.db               — SQLite, bảng weather_daily + weather_hourly (cho bước SQL)

Chạy:  python crawl_weather.py            (mặc định 365 ngày, kết thúc 5 ngày trước hôm nay)
       python crawl_weather.py --days 730
Không cần API key. Open-Meteo Archive có độ trễ ~5 ngày nên end_date lùi 5 ngày.
"""
import argparse
import sqlite3
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests

LAT, LON = 10.8231, 106.6297          # TP.HCM (Quận 1)
TZ = "Asia/Bangkok"                    # UTC+7, Open-Meteo chấp nhận tên này
URL = "https://archive-api.open-meteo.com/v1/archive"

DAILY_VARS = [
    "temperature_2m_max", "temperature_2m_min", "temperature_2m_mean",
    "relative_humidity_2m_max", "relative_humidity_2m_min", "relative_humidity_2m_mean",
    "precipitation_sum", "rain_sum", "precipitation_hours",
    "wind_speed_10m_max", "wind_gusts_10m_max",
]
HOURLY_VARS = ["temperature_2m", "relative_humidity_2m", "precipitation", "wind_speed_10m"]

BASE = Path(__file__).parent
DATA = BASE / "data"


def fetch(start: date, end: date, retries: int = 3) -> dict:
    params = {
        "latitude": LAT, "longitude": LON, "timezone": TZ,
        "start_date": start.isoformat(), "end_date": end.isoformat(),
        "daily": ",".join(DAILY_VARS), "hourly": ",".join(HOURLY_VARS),
        "wind_speed_unit": "kmh",
    }
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(URL, params=params, timeout=60)
            r.raise_for_status()
            return r.json()
        except requests.RequestException as e:
            print(f"  Lần {attempt} lỗi: {e}")
            if attempt == retries:
                raise
            time.sleep(2 * attempt)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=365)
    ap.add_argument("--lag", type=int, default=5, help="lùi end_date bao nhiêu ngày")
    args = ap.parse_args()

    end = date.today() - timedelta(days=args.lag)
    start = end - timedelta(days=args.days - 1)
    print(f"Crawl TP.HCM {start} → {end} ...")
    js = fetch(start, end)

    daily = pd.DataFrame(js["daily"]).rename(columns={"time": "date"})
    daily.columns = [c.replace("_2m", "").replace("_10m", "") for c in daily.columns]
    # temperature_max, temperature_min, temperature_mean, relative_humidity_max, ...
    daily = daily.rename(columns={
        "temperature_max": "temp_max", "temperature_min": "temp_min",
        "temperature_mean": "temp_mean",
        "relative_humidity_max": "rh_max", "relative_humidity_min": "rh_min",
        "relative_humidity_mean": "rh_mean",
        "precipitation_sum": "precip_mm", "rain_sum": "rain_mm",
        "precipitation_hours": "precip_hours",
        "wind_speed_max": "wind_max_kmh", "wind_gusts_max": "gust_max_kmh",
    })

    hourly = pd.DataFrame(js["hourly"]).rename(columns={
        "time": "ts", "temperature_2m": "temp", "relative_humidity_2m": "rh",
        "precipitation": "precip_mm", "wind_speed_10m": "wind_kmh",
    })

    # Làm sạch cơ bản: bỏ ngày thiếu toàn bộ, báo số ô null còn lại
    daily = daily.dropna(how="all", subset=[c for c in daily.columns if c != "date"])
    print(f"  daily : {len(daily)} dòng, null còn lại = {int(daily.isna().sum().sum())}")
    print(f"  hourly: {len(hourly)} dòng, null còn lại = {int(hourly.isna().sum().sum())}")

    DATA.mkdir(exist_ok=True)
    daily.to_csv(DATA / "weather_daily.csv", index=False)
    hourly.to_csv(DATA / "weather_hourly.csv", index=False)

    with sqlite3.connect(BASE / "weather.db") as con:
        daily.to_sql("weather_daily", con, if_exists="replace", index=False)
        hourly.to_sql("weather_hourly", con, if_exists="replace", index=False)
    print("Xong → data/*.csv và weather.db")


if __name__ == "__main__":
    main()