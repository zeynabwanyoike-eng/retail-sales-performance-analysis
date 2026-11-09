
-- 1. Overall sales performance
SELECT
    ROUND(SUM(Revenue), 2) AS total_revenue,
    COUNT(DISTINCT Invoice) AS total_invoices,
    COUNT(DISTINCT "Customer ID") AS unique_customers,
    ROUND(AVG(Revenue), 2) AS average_line_revenue
FROM online_retail;


-- 2. Monthly sales performance
SELECT
    SUBSTR(InvoiceDate, 1, 7) AS sales_month,
    ROUND(SUM(Revenue), 2) AS monthly_revenue,
    COUNT(DISTINCT Invoice) AS invoices
FROM online_retail
GROUP BY sales_month
ORDER BY sales_month;


-- 3. Average order value
SELECT
    ROUND(
        SUM(Revenue) / COUNT(DISTINCT Invoice), 2
    ) AS average_order_value
FROM online_retail;
