-- Revised SQL: RFM Segmentation with dynamic current date
-- We use CURRENT_DATE() to ensure no future dates result in negative recency.

WITH rfm_raw AS (
  SELECT
    user_id,
    -- Recency: Days since the last 'Complete' order compared to today
    DATE_DIFF(CURRENT_DATE(), DATE(MAX(created_at)), DAY) AS recency,
    -- Frequency: Total number of unique orders
    COUNT(DISTINCT order_id) AS frequency,
    -- Monetary: Total sales value from all completed orders
    SUM(sale_price) AS monetary
  FROM
    `bigquery-public-data.thelook_ecommerce.order_items`
  WHERE
    status = 'Complete'
  GROUP BY
    user_id
)
SELECT
  *,
  -- Simple Scoring Logic for Portfolio Demonstration
  CASE 
    WHEN recency <= 30 AND frequency >= 3 THEN 'Champions'
    WHEN recency <= 90 AND frequency >= 2 THEN 'Loyal Customers'
    WHEN recency > 180 AND monetary > 500 THEN 'At Risk (High Value)'
    WHEN recency > 180 AND monetary <= 500 THEN 'Hibernating'
    ELSE 'About to Sleep/Potential'
  END AS segment
FROM
  rfm_raw
ORDER BY 
  monetary DESC;
