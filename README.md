# Customer-Analytics-Dashboard-with-Excel

## Executive Summary
In today’s data-heavy business environment, the ability to pivot from raw transactional records to high-level strategic decisions is a critical competitive advantage. This project features a comprehensive interactive Excel dashboard designed to analyze $29,063,066 in total revenue generated from 452,603 transactions across North American markets.

By synthesizing complex datasets involving thousands of products and nearly 9,000 customers, the solution provides an end-to-end view of organizational health. The dashboard serves as a "single source of truth," allowing stakeholders to monitor Gross Profit ($17,341,934), track order volumes, and evaluate the success of regional operations in the USA, Mexico, and Canada. Through the application of Power Query, Power Pivot, and DAX, the project demonstrates how advanced spreadsheet engineering can deliver the same depth of insight as premium Business Intelligence (BI) platforms at a fraction of the cost.

<img width="1794" height="937" alt="Interactive Customer Dashboard" src="https://github.com/user-attachments/assets/d6b17ebd-2378-4e14-9b4b-a71cdacf7936" />


## Business Problem
The primary challenge addressed by this project was the lack of centralized, dynamic visibility into customer behavior and regional performance. Stakeholders relied on fragmented CSV reports that hindered their ability to identify trends or act on shifting market conditions. To resolve this, a meeting was held to elicit specific business requirements, including metrics, rankings, and segmentations. The objective was to build a tool capable of answering these essential business questions:

* **Performance Tracking:** What is the total Quantity Ordered, Average Order Value, and Gross Profit for any given period?
* **Customer Valuation:** Who are the Top 5 Customers by revenue or quantity, and how can the organization improve retention for these high-value individuals?
* **Demographic Targeting:** How does profit distribution vary across age groups (0-30, 31-50, 50+), gender, and marital status?
* **Geographic Optimization:** Which countries and specific cities represent the most profitable hubs, and which "Bottom 5" areas require operational intervention?
* **Product & Store Analysis:** What is the total count of active products, stores, and customers involved in the $29M revenue cycle?
* **Growth Trends:** What are the Month-over-Month (MoM) % Revenue changes?

## Methodology and Tools
The project utilizes a data pipeline that ensures data integrity, scalability, and performance.

* **Microsoft Excel:** The core environment for analysis and visualization.
8 **Power Query (ETL):** Employed for extracting, transforming, and loading data while automating the cleaning process.
* **Power Pivot (Data Modeling):** Used to build a sophisticated Star Schema and manage massive datasets beyond standard Excel cell limits.
* **DAX (Data Analysis Expressions):** Leveraged to create calculated columns and measures for complex business logic.

### Data Gathering, Cleaning, and Preprocessing
The project began with the extraction of structured data from the corporate data warehouse, saved as CSV files. This included Fact Tables for transactions (2020–2022) and multiple Lookup Tables for customers, products, regions, and stores.

**Power Query** acted as the cleaning engine to prepare this raw data for analysis:

* **Automated Pipeline:** Multiple annual transaction files were appended into a single table. This automation ensures that when new data is dropped into the source folder, a simple "Refresh" updates the entire dashboard.
* **Data Transformation:** A `full_name` column was created by merging first and last names to facilitate customer-level analysis.
* **Data Quality Fixes:** The process involved removing empty rows (top 5 rows of the Store_Lookup), promoting headers, and resolving date-type errors by adjusting the locale settings to English (United States) to match the PC's system time.
* **Redundancy Removal:** Irrelevant fields such as occupation and education were stripped to optimize the data model's size and performance.

<img width="1920" height="1007" alt="Power Query Editor showing the cleaned data pipeline and appended tables" src="https://github.com/user-attachments/assets/0f424c70-d509-4bec-a375-8096c4254233" />


### Data Modelling with Power Pivot
To support fast, cross-functional analysis, a **Star Schema** was implemented within Power Pivot. This architecture places the Fact Table at the center, surrounded by relevant Lookup Tables connected via one-to-many relationships.

* **Bridge Tables:** Because some tables lacked direct join keys, bridge tables were utilized to maintain integrity. For instance, the `Product_Brand_Lookup` was linked to the Fact Table via the `Product_Lookup` using `brand_id` and `product_id`. Similarly, the `Region_Lookup` connected through the `Store_Lookup`.
* **Time Intelligence:** A custom `Calendar_Lookup` table was created and linked to all transaction dates. This is a foundational step for drill-down analysis, allowing stakeholders to view performance by day, weekday, or month.

<img width="1920" height="1005" alt="Data Model Relationship Diagram showing the Star Schema" src="https://github.com/user-attachments/assets/0f5c97b6-4377-4fc6-bebf-23b02c958fdc" />


### Data Analysis Expressions (DAX)
DAX was used to define business logic calculations of the dashboard. Moving beyond basic pivot tables, custom measures were created to provide deeper insights:

* **Core Metrics:** Revenue, COGS, Total Customers, and Average Order Quantity.
* **Advanced Logic:** Time-intelligence formulas were used to calculate Month-over-Month (MoM) % Revenue growth.
* **Demographic Segmentation:** A calculated column for Age was defined using the `DATEDIFF` function, which was then grouped into "Age Buckets" to facilitate the Profit by Age Group visualization.

### Creating the Dashboard
Dashboard design prioritized clarity, professionalism, and user engagement.

#### Canvas and Visual Elements
The dashboard canvas was optimized by removing gridlines and headings to create a focused, software-like interface. A professional color palette was sourced to ensure visual balance, using dark blue for navigation elements and a clean white/grey background for the main visual tiles.

#### Interactive Elements
The interface features dynamic filters for Year, Month, and Country on the left pane. A notable technical feature is the Last Update Date and Time stamp at the top right. This was achieved by creating a blank query in Power Query using M-Code (`DateTime.LocalNow()`) to capture the exact moment of the data refresh, providing users with immediate confidence in the data's recency.

#### Chart Selection
Every visualization was chosen to serve a specific analytical purpose:

* **Dynamic Bar Chart (Top 5 Customers):** This features an "on/off" switch using Form Control option buttons. Users can toggle the ranking by **Quantity Ordered, Total Revenue, or Gross Profit**. This allows the sales team to pivot between volume-based and value-based customer evaluations instantly.
* **Doughnut Charts (Profit by Age Group):** These visualize "parts-of-a-whole" profit distribution. Custom number formatting was applied to show values in millions (e.g., **8.8M for the 0-30 age group**) while displaying the percentage contribution in the center of the doughnut.
* **Map Visualization (Total Profit by Country):** Utilizing Excel's map chart features, this tile provides a geographic overview of the North American market. It uses arrows and callouts to highlight that the **USA drives 61% of total profits**, making the regional dominance immediately apparent to executives.
* **Column Charts (Top/Bottom 5 Cities):** These charts dynamically update based on the country slicer. If a user selects "Mexico," the charts display the most and least profitable cities in that specific region (e.g., **Hidalgo and Merida** as top performers).

<img width="1437" height="538" alt="Dynamic Top 5 Customer toggle and Map Visualization" src="https://github.com/user-attachments/assets/1c536d9a-c970-4907-8b07-61d5e7e71672" />


## Key Metrics (KPIs)
The dashboard tracks the following performance indicators:

* **Total Revenue:** $29,063,066
* **Gross Profit:** $17,341,934
* **Quantity Ordered:** 13,726,397
* **Total Transactions:** 452,603
* **Customer Base:** 8,842
* **Return Volume:** 46,734
* **Product Catalog:** 1,559 active products across 24 stores.

## Insights and Recommendations
The analysis uncovered several critical opportunities for the business:

* **Regional Strategy:** With the USA contributing $10,543,932 (61%). However, Mexico ($5,563,517) represents a significant secondary market (32%).
  * Recommendation: Prioritize supply chain and marketing investments in the USA, particularly in high-performing cities, while investigating the "Bottom 5" cities like Mill Valley and San Francisco to determine if operational costs or low demand are eroding margins.
* **Demographic Focus:** Single customers account for 69% of profit, and those in the 0-30 age group contribute 51% ($8.8M).
  * Recommendation: Marketing campaigns should shift focus toward younger, single demographics. Digital engagement strategies tailored to these segments are likely to yield the highest ROI.
* **Customer Retention:** High-value buyers like Merridee Archuleta and Zelma Pereira were identified as top quantity purchasers.
  * Recommendation: Implement a VIP loyalty program for these top contributors to secure long-term revenue and prevent churn to competitors.

## Conclusion
This project demonstrates that **Advanced Excel** is a high-performance, cost-effective tool for professional Business Intelligence. By transforming nearly half a million rows of raw transactional data into a dynamic, automated dashboard, the solution provides the strategic clarity required to maximize profitability. The value created lies in the transition from **descriptive data** (what happened) to **prescriptive strategy** (what to do next), empowering organizations to make data-driven decisions that optimize North American growth.
