
-- 1. Top 10 products by revenue
SELECT
    StockCode,
    Description,
    SUM(Quantity) AS units_sold,
    ROUND(SUM(Revenue), 2) AS total_revenue
FROM online_retail
GROUP BY StockCode, Description
ORDER BY total_revenue DESC
LIMIT 10;


-- 2. Top 10 products by quantity sold
SELECT
    StockCode,
    Description,
    SUM(Quantity) AS units_sold
FROM online_retail
GROUP BY StockCode, Description
ORDER BY units_sold DESC
LIMIT 10;


-- 3. Products generating the least revenue
SELECT
    StockCode,
    Description,
    SUM(Quantity) AS units_sold,
    ROUND(SUM(Revenue), 2) AS total_revenue
FROM online_retail
GROUP BY StockCode, Description
ORDER BY total_revenue ASC
LIMIT 10;
