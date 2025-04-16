# Superstore Sales Analysis
An Excel-based project analysing sales data to uncover actionable insights and optimise marketing and sales strategies

<img src="assets/img/project3-online-shopping.png" width="500" height="600" />

Image from [storyset](https://storyset.com/search?q=online%20shopping)

> **Note**
>
> This project draws inspiration from Karina Samsonova's data portfolio work featured on her [YouTube channel](https://www.youtube.com/watch?v=U0I3HEnTAWk&ab_channel=KarinaDataScientist). Her comprehensive tutorial on building an interactive Excel dashboard provided exceptional guidance throughout my learning journey. I am profoundly grateful for Karina's expertise and generosity in sharing such valuable educational content.
>
> Original Project:
> 
> Title: Data Analysis Project in Excel - Build Interactive Dashboard. Author: Karina Samsonova Platform: [YouTube](https://www.youtube.com/watch?v=U0I3HEnTAWk&ab_channel=KarinaDataScientist)
>
> This project enhanced my analytical capabilities and expanded my proficiency in data visualisation. Full recognition for the original concept and methodological approach belongs to Karina Samsonova. Her work can be explored further via her [YouTube Channel](https://www.youtube.com/@karinadatascientist) and [LinkedIn](https://www.linkedin.com/in/karina-samsonova/) profile.
<br/>


## Table Of Contents
- [Executive Summary](#executive-summary)
  - [Primary Goal](#primary-goal)
  - [Solution](#solution)
  - [Key Findings](#key-findings)
  - [Recommendations](#recommendations)
- [Introduction](#introduction)
  - [Business Problem](#business-problem)
  - [Goals](#goals)
- [Methodology](#methodology)
  - [Data Source](#data-source)
  - [Tools](#tools)
  - [Data Cleaning](#data-cleaning)
  - [Data Transformation](#data-transformation)
  - [Analysis](#analysis)
    - [Data Visualisation](#data-visualisation)
- [Insights](#insights)
- [Action Plan](#action-plan)	 
- [Excel Report](#excel-report)	


## Executive Summary
### Primary Goal
The objective is to analyse the superstore dataset to uncover actionable insights that drive strategic business decisions and enhance overall performance. By examining key metrics—such as sales trends, customer demographics, and product categories—this analysis aims to identify meaningful patterns, reveal growth opportunities, and detect operational inefficiencies

### Solution
An interactive dashboard will be developed to deliver actionable insights derived from the superstore dataset. This dashboard will focus on analysing key metrics, including sales trends, customer demographics, and product categories. The insights gained will help to identify patterns, uncover growth opportunities, and address inefficiencies. This tool will empower the business team to make informed, data-driven decisions, ultimately improving strategic outcomes and performance.


### Key Findings


### Recommendations


## Introduction
### Business Problem
This analysis aims to address key business challenges through the strategic use of sales data. By leveraging customer segmentation, it identifies high-value consumer groups to refine marketing approaches. It evaluates product performance to optimise stock levels and boost sales, while uncovering regional opportunities through the examination of geographical trends. Additionally, the analysis explores sales patterns and customer behaviour to derive actionable insights that drive growth, improve customer satisfaction, and maximise profitability. The ultimate goal is to establish data-driven strategies that enhance business performance and support sustainable expansion.

### Goals
- Refine marketing strategies by pinpointing high-value consumer groups and tailoring campaigns to their preferences.
- Optimise inventory management through product performance analysis and accurate demand forecasting to maintain efficient stock levels.
- Uncover regional opportunities by identifying geographical trends and improving sales in underperforming areas.
- Maximise profitability by analysing sales patterns and profit margins, enabling focus on high-impact products or categories.

## Methodology
### Data source
[Kaggle](https://www.kaggle.com/datasets/laibaanwer/superstore-sales-dataset) served as the source for the dataset used in this project. Click [here](assets/data/project3_superstore_orders.csv) to access the CSV file directly.

This dataset provides comprehensive information about the sales of various products offered by the store. It includes related details such as geographical data, product categories and subcategories, sales figures, profit margins, and consumer insights.

| Column name | Description | 
| :--- | :--- |
| `order_id` | A unique identifier assigned to each order | 
| `order_date` | The date on which a customer placed their order | 
| `ship_date` | The shipping date corresponding to the order | 
| `ship_mode` | The shipping method chosen for the order | 
| `customer_name` | The customer's full name. | 
| `segment` | The segment type associated with the order. | 
| `state` | The state where the order was placed. | 
| `country` | The country in which the order was placed. | 
| `market` | The market related to the order. Markets include Africa (Africa), Asia and Pacific (APAC), Canada (Canada), Europe, the Middle East, and Africa (EMEA), Europe (EU), Latin America (LATAM), and the United States of America (US). | 
| `region` | The region within the market where the order was originated. | 
| `product_id` | A unique identifier assigned to each product. | 
| `category` | The product's main category. | 
| `sub_category` | The product's subcategory. | 
| `product_name` | The name of the product. | 
| `sales` | The total sales revenue generated for an order, measured in US dollars. | 
| `quantity` | The number of products included in each order. | 
| `discount` | The discount applied to the order. | 
| `profit` | The profit generated from the sale of an order. | 
| `shipping_cost` | The shipping cost incurred for the order. | 
| `order_priority` | The priority level assigned to each order. | 
| `year` | The year in which the order was placed. | 

### Tools
- Excel: To explore, clean, analyse, and visualise the data through a dashboard.


### Data Cleaning and Transformation
After importing the dataset into Excel and converting the file into an Excel Worksheet, I converted the data into an Excel Table to manage and analyse the data easier. The dataset already contained a `year` column, but I considered best to use an Excel formula to extract the year from the `order_date` colum. First, I changed the format of `order_date` column into a Date, then I used an Excel formula to extract the year, formatted the year as a number with no decimals, and named the colum `order_year`.

In most business each customer has a customer ID assigned, but in this dataset that variable is missing. However, using the `customer_name` variable, it was possible to identify unique customers using a conditional Excel formula to count each time a new customer is added to the database. This new variable was named `customer_unique` 

Another variable to count was the unique number of orders. In this case, the database has a `order_id` column, so I used a similar conditional formula to count the unique number of `order_id`s, the new variable was name `order_unique`. 


### Data Analysis
The first step was to estimte four KPIs using Pivot Tables: `Total Sales (USD)`, the `Number of Unique Orders`, and the `Number of Unique Customers`. Then, I created a calculated field to estimate the `Average Sales per Order (USD)`, leaving the result with two decimals.
  
Using pivot tables I calculate the total sales by quarter and total sales by year using the `order_year` variable to analyse the Sales Trend over time.  

To analyse regional performances, I use a Pivot Table to see the Total Sales by Market. Then with another pivot table, I selected the Top 10 Countries Leading in Total Sales (USD).

For customers insights, using a pivot table I first estimate the Total Sales (USD) for each of the three segments (Consumer, Corportate, and Home Office). After that, I analyse the Customer Preferences for Shipping Modes in terms of % total orders.

Finally, to get some products insights, I used pivot tables to see the Total Sales (USD) by Product Category and Subcategory,  and a table with the Total Sales (USD) of the Top 10 Best-Selling Products in USD.  

### Excel formulas
| Formula | Description | 
| :--- | :--- |
| `=YEAR([@[order_date]])` | Extract the Year from the order date |
| `=IF(COUNTIF(F$2:F2,F2)=1,1,0) ` | Count each unique customer. If the user is unique, it assigns a 1 value, if the customer is repeated, it assigs a 0.  |
| `=IF(COUNTIF(A$2:A2,A2)=1,1,0) ` | Count each unique order_id. If an order_id is unique, it assigns a 1 value, if the order_id is repeated, it assigs a 0. |
| `Avg Sales per Order = sales/ order_unique ` | A calculated field to estimate the average sales per order  |

#### Data Visualisation

Then, I added a slicer in the `order_year` to filter the Pivot Table and the final dashboard.



## Insights

## Action plan


## Excel Report
To review the analysis in detail, you can download the Excel Report [here]().
