
-- 1. Total sales revenue
SELECT ROUND(SUM(Revenue), 2) AS total_revenue
FROM online_retail;

-- 2. Total transactions
SELECT COUNT(DISTINCT Invoice) AS total_invoices
FROM online_retail;

-- 3. Monthly revenue
SELECT
    strftime('%Y-%m', InvoiceDate) AS month,
    ROUND(SUM(Revenue), 2) AS monthly_revenue
FROM online_retail
GROUP BY month
ORDER BY month;

-- 4. Top 10 best-selling products
SELECT
    StockCode,
    Description,
    SUM(Quantity) AS units_sold,
    ROUND(SUM(Revenue), 2) AS revenue
FROM online_retail
GROUP BY StockCode, Description
ORDER BY revenue DESC
LIMIT 10;

-- 5. Top 10 customers
SELECT
    "Customer ID",
    ROUND(SUM(Revenue), 2) AS customer_revenue
FROM online_retail
WHERE "Customer ID" IS NOT NULL
GROUP BY "Customer ID"
ORDER BY customer_revenue DESC
LIMIT 10;

-- 6. Revenue by country
SELECT
    Country,
    ROUND(SUM(Revenue), 2) AS country_revenue
FROM online_retail
GROUP BY Country
ORDER BY country_revenue DESC;
