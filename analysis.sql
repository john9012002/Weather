-- =====================================================================
-- Bước 2 — Phân tích bằng SQL (SQLite ≥ 3.25; chạy được gần như nguyên
-- vẹn trên PostgreSQL — xem ghi chú [PG] ở các dòng khác biệt).
-- Nguồn: bảng weather_daily, weather_hourly do crawl_weather.py tạo.
--
-- Quy ước:
--   * Mùa mưa TP.HCM: tháng 5 → 11 ; mùa khô: tháng 12 → 4.
--   * "Ngày có mưa" (wet day): lượng mưa ≥ 1 mm (ngưỡng WMO).
--   * Ban ngày: 06:00–17:59 ; ban đêm: 18:00–05:59 (cùng ngày dương lịch).
--   * Tăng vọt đột ngột (spike): temp_max hôm nay cao hơn hôm qua ≥ 2 °C
--     HOẶC cao hơn trung bình 7 ngày TRƯỚC đó ≥ 2 °C.
-- =====================================================================

-- ---------------------------------------------------------------------
-- V1. Bảng ngày đã làm giàu: mùa, MA7, LAG/LEAD, cờ spike
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_daily_enriched;
CREATE VIEW v_daily_enriched AS
WITH base AS (
    SELECT
        date,
        CAST(strftime('%m', date) AS INTEGER)            AS month,      -- [PG] EXTRACT(MONTH FROM date::date)
        CASE WHEN CAST(strftime('%m', date) AS INTEGER) BETWEEN 5 AND 11
             THEN 'Mùa mưa' ELSE 'Mùa khô' END            AS season,
        temp_max, temp_min, temp_mean,
        ROUND(temp_max - temp_min, 2)                     AS diurnal_range,
        rh_mean, rh_max, rh_min,
        precip_mm, precip_hours, wind_max_kmh,
        CASE WHEN precip_mm >= 1 THEN 1 ELSE 0 END        AS is_wet
    FROM weather_daily
),
win AS (
    SELECT
        b.*,
        -- Trung bình động 7 ngày (gồm hôm nay + 6 ngày trước)
        ROUND(AVG(temp_mean) OVER w7, 2)                  AS temp_mean_ma7,
        ROUND(AVG(temp_max)  OVER w7, 2)                  AS temp_max_ma7,
        -- Trung bình 7 ngày TRƯỚC (không gồm hôm nay) — mốc so sánh spike
        ROUND(AVG(temp_max) OVER (ORDER BY date
              ROWS BETWEEN 7 PRECEDING AND 1 PRECEDING), 2) AS temp_max_prev7,
        COUNT(*) OVER w7                                  AS ma7_n,     -- < 7 ở 6 ngày đầu
        LAG(temp_max)  OVER (ORDER BY date)               AS temp_max_prev_day,
        LEAD(temp_max) OVER (ORDER BY date)               AS temp_max_next_day,
        LEAD(is_wet)   OVER (ORDER BY date)               AS is_wet_next_day
    FROM base b
    WINDOW w7 AS (ORDER BY date ROWS BETWEEN 6 PRECEDING AND CURRENT ROW)
)
SELECT
    w.*,
    ROUND(temp_max - temp_max_prev_day, 2)                AS delta_vs_yesterday,
    ROUND(temp_max - temp_max_prev7, 2)                   AS delta_vs_prev7,
    CASE WHEN temp_max - temp_max_prev_day >= 2
           OR temp_max - temp_max_prev7    >= 2
         THEN 1 ELSE 0 END                                AS is_spike,
    -- Spike "chớp nhoáng": tăng vọt rồi hôm sau hạ lại ≥ 1.5 °C
    CASE WHEN (temp_max - temp_max_prev_day >= 2 OR temp_max - temp_max_prev7 >= 2)
          AND temp_max - temp_max_next_day >= 1.5
         THEN 1 ELSE 0 END                                AS is_short_spike
FROM win w;


-- ---------------------------------------------------------------------
-- V2. Ngày cực đoan: top 10% nhiệt độ / độ ẩm (PERCENT_RANK)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_extreme_days;
CREATE VIEW v_extreme_days AS
WITH ranked AS (
    SELECT date, month, season, temp_max, temp_min, rh_mean, rh_min, precip_mm,
           PERCENT_RANK() OVER (ORDER BY temp_max)      AS pr_temp_hot,
           PERCENT_RANK() OVER (ORDER BY temp_min DESC) AS pr_temp_cool,
           PERCENT_RANK() OVER (ORDER BY rh_mean)       AS pr_rh_humid,
           PERCENT_RANK() OVER (ORDER BY rh_mean DESC)  AS pr_rh_dry
    FROM v_daily_enriched
)
SELECT *,
       TRIM(
         CASE WHEN pr_temp_hot  >= 0.9 THEN 'Nóng cực đoan; '  ELSE '' END ||
         CASE WHEN pr_temp_cool >= 0.9 THEN 'Mát cực đoan; '   ELSE '' END ||
         CASE WHEN pr_rh_humid  >= 0.9 THEN 'Ẩm cực đoan; '    ELSE '' END ||
         CASE WHEN pr_rh_dry    >= 0.9 THEN 'Khô cực đoan; '   ELSE '' END
       ) AS extreme_tags
FROM ranked
WHERE pr_temp_hot >= 0.9 OR pr_temp_cool >= 0.9
   OR pr_rh_humid >= 0.9 OR pr_rh_dry >= 0.9;


-- ---------------------------------------------------------------------
-- V3. Chênh lệch ngày – đêm theo ngày (từ dữ liệu giờ)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_day_night;
CREATE VIEW v_day_night AS
SELECT
    substr(ts, 1, 10)                                              AS date,   -- [PG] ts::date
    ROUND(AVG(CASE WHEN CAST(substr(ts,12,2) AS INTEGER) BETWEEN 6 AND 17
                   THEN temp END), 2)                              AS temp_day_avg,
    ROUND(AVG(CASE WHEN CAST(substr(ts,12,2) AS INTEGER) NOT BETWEEN 6 AND 17
                   THEN temp END), 2)                              AS temp_night_avg,
    ROUND(AVG(CASE WHEN CAST(substr(ts,12,2) AS INTEGER) BETWEEN 6 AND 17
                   THEN temp END)
        - AVG(CASE WHEN CAST(substr(ts,12,2) AS INTEGER) NOT BETWEEN 6 AND 17
                   THEN temp END), 2)                              AS day_night_gap
FROM weather_hourly
GROUP BY substr(ts, 1, 10);


-- ---------------------------------------------------------------------
-- V4. Tổng hợp theo tháng (dùng cho Dashboard Excel)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_monthly;
CREATE VIEW v_monthly AS
SELECT
    substr(e.date, 1, 7)                       AS year_month,
    e.month, e.season,
    COUNT(*)                                   AS n_days,
    ROUND(AVG(e.temp_max), 2)                  AS avg_temp_max,
    ROUND(AVG(e.temp_min), 2)                  AS avg_temp_min,
    ROUND(AVG(e.diurnal_range), 2)             AS avg_diurnal_range,
    ROUND(AVG(d.temp_day_avg), 2)              AS avg_temp_day,
    ROUND(AVG(d.temp_night_avg), 2)            AS avg_temp_night,
    ROUND(AVG(e.rh_mean), 2)                   AS avg_rh,
    ROUND(SUM(e.precip_mm), 1)                 AS total_rain_mm,
    SUM(e.is_wet)                              AS wet_days,
    ROUND(1.0 * SUM(e.is_wet) / COUNT(*), 3)   AS rain_probability,
    SUM(e.is_spike)                            AS spike_days
FROM v_daily_enriched e
LEFT JOIN v_day_night d ON d.date = e.date
GROUP BY substr(e.date, 1, 7), e.month, e.season
ORDER BY year_month;


-- ---------------------------------------------------------------------
-- V5. "Dự báo" khả năng mưa theo mùa
--   * P(mưa)                 : tần suất ngày mưa trong mùa
--   * P(mai mưa | nay mưa)   : chuỗi Markov 1 bậc, dùng LEAD()
--   * P(mai mưa | nay khô)
-- ---------------------------------------------------------------------
DROP VIEW IF EXISTS v_rain_forecast_season;
CREATE VIEW v_rain_forecast_season AS
SELECT
    season,
    COUNT(*)                                                     AS n_days,
    ROUND(AVG(is_wet), 3)                                        AS p_rain,
    ROUND(1.0 * SUM(CASE WHEN is_wet = 1 AND is_wet_next_day = 1 THEN 1 ELSE 0 END)
        / NULLIF(SUM(CASE WHEN is_wet = 1 AND is_wet_next_day IS NOT NULL THEN 1 ELSE 0 END), 0), 3)
                                                                 AS p_rain_given_rain_today,
    ROUND(1.0 * SUM(CASE WHEN is_wet = 0 AND is_wet_next_day = 1 THEN 1 ELSE 0 END)
        / NULLIF(SUM(CASE WHEN is_wet = 0 AND is_wet_next_day IS NOT NULL THEN 1 ELSE 0 END), 0), 3)
                                                                 AS p_rain_given_dry_today,
    ROUND(AVG(temp_max), 2)                                      AS avg_temp_max,
    ROUND(AVG(rh_mean), 2)                                       AS avg_rh,
    ROUND(SUM(precip_mm), 1)                                     AS total_rain_mm
FROM v_daily_enriched
GROUP BY season;


-- ---------------------------------------------------------------------
-- Q. Câu truy vấn kiểm tra nhanh (chạy tay trong DB Browser / DBeaver)
-- ---------------------------------------------------------------------
-- Top 10 ngày tăng nhiệt đột ngột nhất
-- SELECT date, temp_max_prev_day, temp_max, delta_vs_yesterday, delta_vs_prev7
-- FROM v_daily_enriched WHERE is_spike = 1
-- ORDER BY MAX(delta_vs_yesterday, delta_vs_prev7) DESC LIMIT 10;   -- [PG] GREATEST(...)

-- Chuỗi ngày mưa liên tiếp dài nhất (gaps-and-islands)
-- WITH g AS (SELECT date, is_wet,
--            ROW_NUMBER() OVER (ORDER BY date)
--          - ROW_NUMBER() OVER (PARTITION BY is_wet ORDER BY date) AS grp
--            FROM v_daily_enriched)
-- SELECT MIN(date) AS start_date, MAX(date) AS end_date, COUNT(*) AS streak
-- FROM g WHERE is_wet = 1 GROUP BY grp ORDER BY streak DESC LIMIT 5;