
-- 1. Top 10 customers by revenue
SELECT
    "Customer ID",
    ROUND(SUM(Revenue), 2) AS total_spent,
    COUNT(DISTINCT Invoice) AS total_orders
FROM online_retail
WHERE "Customer ID" IS NOT NULL
GROUP BY "Customer ID"
ORDER BY total_spent DESC
LIMIT 10;


-- 2. Customer purchase frequency
SELECT
    "Customer ID",
    COUNT(DISTINCT Invoice) AS purchase_count,
    ROUND(SUM(Revenue), 2) AS total_spent
FROM online_retail
WHERE "Customer ID" IS NOT NULL
GROUP BY "Customer ID"
ORDER BY purchase_count DESC
LIMIT 10;


-- 3. Average customer spending
SELECT
    ROUND(AVG(customer_revenue), 2) AS average_customer_spending
FROM (
    SELECT
        "Customer ID",
        SUM(Revenue) AS customer_revenue
    FROM online_retail
    WHERE "Customer ID" IS NOT NULL
    GROUP BY "Customer ID"
);
