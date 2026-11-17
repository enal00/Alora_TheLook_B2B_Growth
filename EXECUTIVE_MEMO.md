# Executive Portfolio: Data-Driven B2B Growth at The Look

## 1. Business Understanding
**Objective:** Optimize B2B lead acquisition and operational efficiency.
**Challenge:** High customer acquisition cost (CAC) and stagnating growth in key business units. 
**Goal:** Identify high-value segments and develop actionable growth strategies using RFM (Recency, Frequency, Monetary) and Cohort Analysis.

## 2. Data Source
- **Primary Source:** BigQuery Public Data (`thelook_ecommerce`).
- **Nature of Data:** Real-world e-commerce transactional data, including users, orders, products, and operational events.

## 3. Cloud Environment & Tech Stack
- **Cloud Infrastructure:** Google BigQuery (Enterprise-grade data warehousing).
- **Languages/Tools:** SQL (Standard SQL for data extraction & cleaning), Python (Pandas/Seaborn/Matplotlib for advanced analytics & visual storytelling).
- **Environment:** Local VS Code with managed data pipeline.

## 4. Data Cleaning
- **Anomaly Handling:** Addressed negative `recency` values resulting from live-dataset date offsets by capping to 0 (representing 'today').
- **Financial Integrity:** Filtered out 'Cancelled' and 'Returned' orders to ensure cashflow calculations reflect real operational revenue.

## 5. Data Extraction & Feature Engineering
- **RFM Logic:** Engineered metrics from raw logs:
    - **Recency:** Days since the last 'Complete' transaction.
    - **Frequency:** Total unique orders per user.
    - **Monetary:** Total lifecycle sales value.
- **Segmentation:** Applied business-logic thresholds to categorize users into 'Champions', 'Loyal', 'At Risk (High Value)', and 'Hibernating'.

## 6. Key Findings
- **Market Asymmetry:** Brazil exhibits 49% higher AOV ($153.22) compared to the global average, signaling a high-ROI expansion opportunity.
- **Retention Maturity:** Analysis of Cohort Matrix shows a significant improvement in retention trends from Mar '22 (20% monthly drop) to Oct '22 (only 5% drop), indicating strong product-market fit improvements.
- **The "Whale" Risk:** Identified 37 'At Risk (High Value)' clients with high monetary stakes ($705 avg) who are currently disengaged.

## 7. Dashboard (Visual Summary)
- **Revenue Heatmap:** Highlights revenue concentration across segments.
- **Customer Health (Scatter):** Visualizes the gap between active frequency and recency.
- **Retention Cohort:** Tracks user stickiness over a 12-month lifecycle.

## 8. Insights & Strategic Recommendations
- **B2B Expansion:** Prioritize 'Outerwear & Coats' for the Brazilian market to capitalize on higher AOV.
- **Retention Engine:** Deploy a 'Priority Concierge' program for the 37 identified 'At Risk' high-value clients to secure revenue before churn.
- **Operational Efficiency:** Use the improved cohort retention as a benchmark to audit and replicate successful marketing tactics from Q4 2022 across all active business units.
