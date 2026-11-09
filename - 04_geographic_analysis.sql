
-- 1. Revenue by country
SELECT
    Country,
    ROUND(SUM(Revenue), 2) AS total_revenue
FROM online_retail
GROUP BY Country
ORDER BY total_revenue DESC;


-- 2. Top 10 countries by number of orders
SELECT
    Country,
    COUNT(DISTINCT Invoice) AS total_orders
FROM online_retail
GROUP BY Country
ORDER BY total_orders DESC
LIMIT 10;


-- 3. Average revenue per invoice by country
SELECT
    Country,
    ROUND(
        SUM(Revenue) / COUNT(DISTINCT Invoice), 2
    ) AS average_order_value
FROM online_retail
GROUP BY Country
ORDER BY average_order_value DESC;
