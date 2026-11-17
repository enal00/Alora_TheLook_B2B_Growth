-- Query: Cross-Category Affinity for B2B Cross-Selling
-- This query helps identify products that are frequently bought together by the same users.
-- Essential for developing cross-selling strategies and B2B growth recommendations.

WITH user_orders AS (
  SELECT 
    oi.user_id,
    p.category,
    COUNT(*) as item_count
  FROM `bigquery-public-data.thelook_ecommerce.order_items` oi
  JOIN `bigquery-public-data.thelook_ecommerce.products` p ON oi.product_id = p.id
  WHERE oi.status = 'Complete'
  GROUP BY 1, 2
)
SELECT 
    a.category AS product_a,
    b.category AS product_b,
    COUNT(DISTINCT a.user_id) AS common_users
FROM user_orders a
JOIN user_orders b ON a.user_id = b.id AND a.category < b.category
GROUP BY 1, 2
ORDER BY common_users DESC
LIMIT 15;
