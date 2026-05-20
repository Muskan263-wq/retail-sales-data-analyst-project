-- ============================================================
--  RETAIL SALES DATA ANALYST PROJECT — SQL SCRIPT
--  Level: Intermediate | Tools: SQLite / PostgreSQL / MySQL
-- ============================================================
-- Run these queries after importing retail_sales_analysis.xlsx
-- into your database as a table called "sales"
-- ============================================================

-- ─────────────────────────────────────────────────────────────
-- SECTION 0: TABLE SETUP (SQLite compatible)
-- ─────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS sales (
    Order_ID        TEXT,
    Order_Date      DATE,
    Ship_Date       DATE,
    Customer_ID     TEXT,
    Customer_Name   TEXT,
    Segment         TEXT,
    Region          TEXT,
    State           TEXT,
    Category        TEXT,
    Sub_Category    TEXT,
    Product_Name    TEXT,
    Sales           REAL,
    Quantity        INTEGER,
    Discount        REAL,
    Profit          REAL,
    Profit_Margin   REAL
);


-- ─────────────────────────────────────────────────────────────
-- SECTION 1: BASIC EXPLORATION
-- ─────────────────────────────────────────────────────────────

-- 1.1 Quick overview
SELECT COUNT(*)          AS total_orders,
       COUNT(DISTINCT Customer_ID) AS unique_customers,
       MIN(Order_Date)   AS first_order,
       MAX(Order_Date)   AS last_order,
       ROUND(SUM(Sales),2)  AS total_revenue,
       ROUND(SUM(Profit),2) AS total_profit
FROM sales;

-- 1.2 Null / missing value check
SELECT
    SUM(CASE WHEN Order_ID    IS NULL THEN 1 ELSE 0 END) AS null_order_id,
    SUM(CASE WHEN Sales       IS NULL THEN 1 ELSE 0 END) AS null_sales,
    SUM(CASE WHEN Profit      IS NULL THEN 1 ELSE 0 END) AS null_profit,
    SUM(CASE WHEN Customer_ID IS NULL THEN 1 ELSE 0 END) AS null_customer
FROM sales;

-- 1.3 Duplicate order check
SELECT Order_ID, COUNT(*) AS cnt
FROM sales
GROUP BY Order_ID
HAVING COUNT(*) > 1;


-- ─────────────────────────────────────────────────────────────
-- SECTION 2: AGGREGATION & GROUPING
-- ─────────────────────────────────────────────────────────────

-- 2.1 Sales & Profit by Region (sorted best to worst)
SELECT
    Region,
    COUNT(*)                        AS orders,
    ROUND(SUM(Sales),   2)          AS total_sales,
    ROUND(SUM(Profit),  2)          AS total_profit,
    ROUND(AVG(Profit_Margin)*100,1) AS avg_margin_pct
FROM sales
GROUP BY Region
ORDER BY total_sales DESC;

-- 2.2 Sales by Category + Sub-Category
SELECT
    Category,
    Sub_Category,
    ROUND(SUM(Sales),  2)  AS total_sales,
    ROUND(SUM(Profit), 2)  AS total_profit,
    SUM(Quantity)           AS units_sold,
    ROUND(AVG(Discount)*100,1) AS avg_discount_pct
FROM sales
GROUP BY Category, Sub_Category
ORDER BY Category, total_sales DESC;

-- 2.3 Monthly revenue trend
SELECT
    STRFTIME('%Y', Order_Date) AS year,
    STRFTIME('%m', Order_Date) AS month,
    ROUND(SUM(Sales), 2)       AS monthly_sales,
    ROUND(SUM(Profit),2)       AS monthly_profit,
    COUNT(*)                   AS orders
FROM sales
GROUP BY year, month
ORDER BY year, month;

-- 2.4 Segment performance
SELECT
    Segment,
    ROUND(SUM(Sales),2)           AS total_sales,
    ROUND(AVG(Sales),2)           AS avg_order_value,
    ROUND(SUM(Profit),2)          AS total_profit,
    ROUND(AVG(Discount)*100,1)    AS avg_discount_pct
FROM sales
GROUP BY Segment
ORDER BY total_sales DESC;


-- ─────────────────────────────────────────────────────────────
-- SECTION 3: FILTERING & CONDITIONAL ANALYSIS
-- ─────────────────────────────────────────────────────────────

-- 3.1 High-value orders (Sales > $1,000)
SELECT Order_ID, Customer_Name, Region, Category,
       ROUND(Sales,2) AS sales, ROUND(Profit,2) AS profit
FROM sales
WHERE Sales > 1000
ORDER BY Sales DESC
LIMIT 20;

-- 3.2 Loss-making orders (Profit < 0)
SELECT Order_ID, Customer_Name, Sub_Category,
       ROUND(Sales,2) AS sales, ROUND(Profit,2) AS profit,
       ROUND(Discount*100,0) AS discount_pct
FROM sales
WHERE Profit < 0
ORDER BY Profit ASC
LIMIT 20;

-- 3.3 Impact of discount on profitability
SELECT
    CASE
        WHEN Discount = 0          THEN '0% — No Discount'
        WHEN Discount <= 0.10      THEN '1–10%'
        WHEN Discount <= 0.20      THEN '11–20%'
        ELSE '20%+ — Heavy Discount'
    END AS discount_bucket,
    COUNT(*)                        AS orders,
    ROUND(SUM(Sales),2)             AS total_sales,
    ROUND(AVG(Profit_Margin)*100,1) AS avg_margin_pct,
    ROUND(SUM(Profit),2)            AS total_profit
FROM sales
GROUP BY discount_bucket
ORDER BY avg_margin_pct DESC;

-- 3.4 Best performing states
SELECT
    State,
    Region,
    ROUND(SUM(Sales),2)  AS total_sales,
    ROUND(SUM(Profit),2) AS total_profit,
    COUNT(*)             AS orders
FROM sales
GROUP BY State, Region
ORDER BY total_sales DESC
LIMIT 10;


-- ─────────────────────────────────────────────────────────────
-- SECTION 4: SUBQUERIES
-- ─────────────────────────────────────────────────────────────

-- 4.1 Customers who spent more than average
SELECT Customer_ID, Customer_Name, Segment,
       ROUND(SUM(Sales),2) AS total_spent
FROM sales
GROUP BY Customer_ID, Customer_Name, Segment
HAVING SUM(Sales) > (
    SELECT AVG(customer_sales)
    FROM (
        SELECT Customer_ID, SUM(Sales) AS customer_sales
        FROM sales
        GROUP BY Customer_ID
    ) sub
)
ORDER BY total_spent DESC
LIMIT 15;

-- 4.2 Products below average profit margin in their category
SELECT s.Sub_Category, s.Product_Name,
       ROUND(AVG(s.Profit_Margin)*100,1)  AS product_avg_margin,
       ROUND(cat_avg.avg_margin*100,1)     AS category_avg_margin
FROM sales s
JOIN (
    SELECT Sub_Category, AVG(Profit_Margin) AS avg_margin
    FROM sales
    GROUP BY Sub_Category
) cat_avg ON s.Sub_Category = cat_avg.Sub_Category
GROUP BY s.Sub_Category, s.Product_Name, cat_avg.avg_margin
HAVING AVG(s.Profit_Margin) < cat_avg.avg_margin
ORDER BY s.Sub_Category, product_avg_margin;


-- ─────────────────────────────────────────────────────────────
-- SECTION 5: CTEs (Common Table Expressions)
-- ─────────────────────────────────────────────────────────────

-- 5.1 Year-over-Year comparison
WITH yearly AS (
    SELECT
        STRFTIME('%Y', Order_Date) AS yr,
        ROUND(SUM(Sales),2)        AS total_sales,
        ROUND(SUM(Profit),2)       AS total_profit
    FROM sales
    GROUP BY yr
)
SELECT
    a.yr                                                          AS year,
    a.total_sales,
    b.total_sales                                                 AS prev_year_sales,
    ROUND((a.total_sales - b.total_sales) / b.total_sales * 100, 1) AS sales_growth_pct,
    a.total_profit,
    ROUND((a.total_profit - b.total_profit) / ABS(b.total_profit) * 100, 1) AS profit_growth_pct
FROM yearly a
LEFT JOIN yearly b ON CAST(a.yr AS INTEGER) = CAST(b.yr AS INTEGER) + 1
ORDER BY a.yr;

-- 5.2 Top 3 products per category by sales (CTE + ROW_NUMBER)
WITH ranked AS (
    SELECT
        Category,
        Product_Name,
        ROUND(SUM(Sales),2) AS total_sales,
        ROW_NUMBER() OVER (
            PARTITION BY Category ORDER BY SUM(Sales) DESC
        ) AS rank_in_cat
    FROM sales
    GROUP BY Category, Product_Name
)
SELECT Category, rank_in_cat AS rank, Product_Name, total_sales
FROM ranked
WHERE rank_in_cat <= 3
ORDER BY Category, rank_in_cat;

-- 5.3 Running total of monthly sales (cumulative)
WITH monthly AS (
    SELECT
        STRFTIME('%Y-%m', Order_Date) AS ym,
        ROUND(SUM(Sales),2)           AS monthly_sales
    FROM sales
    GROUP BY ym
)
SELECT
    ym,
    monthly_sales,
    ROUND(SUM(monthly_sales) OVER (ORDER BY ym ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW), 2)
        AS cumulative_sales
FROM monthly
ORDER BY ym;


-- ─────────────────────────────────────────────────────────────
-- SECTION 6: WINDOW FUNCTIONS
-- ─────────────────────────────────────────────────────────────

-- 6.1 Sales rank per region
SELECT
    Customer_Name,
    Region,
    ROUND(SUM(Sales),2) AS total_sales,
    RANK() OVER (
        PARTITION BY Region ORDER BY SUM(Sales) DESC
    ) AS region_rank
FROM sales
GROUP BY Customer_Name, Region
ORDER BY Region, region_rank
LIMIT 20;

-- 6.2 Month-over-month change
WITH monthly AS (
    SELECT
        STRFTIME('%Y-%m', Order_Date) AS ym,
        ROUND(SUM(Sales),2)           AS monthly_sales
    FROM sales GROUP BY ym
)
SELECT
    ym,
    monthly_sales,
    LAG(monthly_sales) OVER (ORDER BY ym)  AS prev_month_sales,
    ROUND(monthly_sales - LAG(monthly_sales) OVER (ORDER BY ym), 2)  AS mom_change
FROM monthly
ORDER BY ym;

-- 6.3 Percent of total sales per category
SELECT
    Category,
    ROUND(SUM(Sales),2) AS cat_sales,
    ROUND(SUM(Sales) * 100.0 / SUM(SUM(Sales)) OVER (), 1) AS pct_of_total
FROM sales
GROUP BY Category
ORDER BY cat_sales DESC;

-- 6.4 Customer purchase frequency buckets
WITH cust_orders AS (
    SELECT Customer_ID, Customer_Name, COUNT(*) AS order_count
    FROM sales GROUP BY Customer_ID, Customer_Name
)
SELECT
    CASE
        WHEN order_count = 1 THEN 'One-time'
        WHEN order_count <= 3 THEN 'Occasional (2–3)'
        WHEN order_count <= 7 THEN 'Regular (4–7)'
        ELSE 'Loyal (8+)'
    END AS customer_segment,
    COUNT(*)          AS num_customers,
    AVG(order_count)  AS avg_orders
FROM cust_orders
GROUP BY customer_segment
ORDER BY avg_orders DESC;


-- ─────────────────────────────────────────────────────────────
-- SECTION 7: VIEWS (reusable analytical layers)
-- ─────────────────────────────────────────────────────────────

-- 7.1 Executive summary view
CREATE VIEW IF NOT EXISTS v_executive_summary AS
SELECT
    Region,
    Category,
    STRFTIME('%Y', Order_Date)  AS year,
    ROUND(SUM(Sales),2)         AS total_sales,
    ROUND(SUM(Profit),2)        AS total_profit,
    ROUND(AVG(Profit_Margin)*100,1) AS avg_margin_pct,
    COUNT(*)                    AS order_count
FROM sales
GROUP BY Region, Category, year;

SELECT * FROM v_executive_summary ORDER BY year, total_sales DESC;

-- 7.2 Customer lifetime value view
CREATE VIEW IF NOT EXISTS v_customer_ltv AS
SELECT
    Customer_ID,
    Customer_Name,
    Segment,
    COUNT(*)                        AS total_orders,
    ROUND(SUM(Sales),2)             AS lifetime_sales,
    ROUND(SUM(Profit),2)            AS lifetime_profit,
    ROUND(AVG(Sales),2)             AS avg_order_value,
    MIN(Order_Date)                 AS first_order_date,
    MAX(Order_Date)                 AS last_order_date
FROM sales
GROUP BY Customer_ID, Customer_Name, Segment;

SELECT * FROM v_customer_ltv ORDER BY lifetime_sales DESC LIMIT 20;


-- ─────────────────────────────────────────────────────────────
-- SECTION 8: INTERVIEW-READY QUERIES
-- ─────────────────────────────────────────────────────────────

-- Q: Which sub-category has the highest loss?
SELECT Sub_Category, ROUND(SUM(Profit),2) AS total_profit
FROM sales GROUP BY Sub_Category ORDER BY total_profit ASC LIMIT 5;

-- Q: Find the month with the highest sales each year
WITH monthly AS (
    SELECT STRFTIME('%Y',Order_Date) AS yr, STRFTIME('%m',Order_Date) AS mo,
           ROUND(SUM(Sales),2) AS sales,
           RANK() OVER (PARTITION BY STRFTIME('%Y',Order_Date)
                        ORDER BY SUM(Sales) DESC) AS rnk
    FROM sales GROUP BY yr, mo
)
SELECT yr AS year, mo AS best_month, sales AS peak_sales FROM monthly WHERE rnk=1;

-- Q: Average days to ship by region
SELECT Region,
       ROUND(AVG(JULIANDAY(Ship_Date) - JULIANDAY(Order_Date)), 1) AS avg_ship_days
FROM sales
GROUP BY Region ORDER BY avg_ship_days;

-- END OF SCRIPT ──────────────────────────────────────────────
