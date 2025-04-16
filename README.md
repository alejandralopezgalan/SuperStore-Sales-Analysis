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
Analyse the superstore dataset to uncover actionable insights that can drive strategic business decisions and improve overall performance. By examining key metrics, such as sales trends, customer demographics, and product categories, we can identify meaningful patterns, highlight growth opportunities, and detect operational inefficiencies.

### Solution



### Key Findings


### Recommendations

## Introduction

### Business Problem
This analysis seeks to tackle critical business challenges by utilising sales data to refine marketing and sales strategies. It explores customer segmentation to identify high-value consumer groups, evaluates product performance to optimise stock levels and increase sales, and examines geographical trends to uncover regional opportunities. Furthermore, it investigates profit margins, sales patterns, and customer behaviour to extract actionable insights that can drive growth, enhance customer satisfaction, and maximise profitability. The ultimate goal is to develop data-driven strategies for business improvement and sustainable expansion.

### Goals
1. Enhance marketing strategies by identifying high-value consumer groups and tailoring campaigns to their preferences.
2. Optimise inventory management by analysing product performance and predicting demand for improved stock levels.
3. Identify regional opportunities through geographical trends to increase sales in underperforming areas.
4. Maximise profitability by evaluating sales patterns and profit margins to focus efforts on high-impact products or categories.

## Methodology
### Data source
[Kaggle](https://www.kaggle.com/datasets/laibaanwer/superstore-sales-dataset) served as the source for the dataset used in this project. Click [here](assets/data/project3_superstore_orders.csv) to view the CSV file directly.

This dataset encompasses information on the sales of various products offered by the store, including related details such as geographical data, product categories and subcategories, sales figures, profit margins, and consumer  insights.

| Column name | Description | 
| :--- | :--- |
| `order_id` | A unique identifier for each order. | 
| `order_date` | The date when a customer placed an order. | 
| `ship_date` | The shipping date for an order. | 
| `ship_mode` | The mode of shipping for each order. | 
| `customer_name` | The full name of the customer. | 
| `segment` | The segment type associated with the order. | 
| `state` | The state where the order was placed. | 
| `country` | The country where the order was placed. | 
| `market` | The market related to the order. África = Africa, APAC = Asia and Pacific, Canada = Canada, EMEA = Europe, the Middle East, and Africa, EU = Europe, LATAM = Latin America, US = United States of America. | 
| `region` | The region within the market where the order was placed. | 
| `product_id` | A unique identifier for each product. | 
| `category` | The product's category. | 
| `sub_category` | The product's subcategory. | 
| `product_name` | The name of the product. | 
| `sales` | Total sales in US dollars for an order. | 
| `quantity` | The quantity of products in each order. | 
| `discount` | The discount applied to each order. | 
| `profit` | The profit generated after the sale of each order. | 
| `shipping_cost` | The shipping cost for every order. | 
| `order_priority` | The priority level assigned to each order. | 
| `year` | The year in which the order was placed. | 


### Tools
- Excel: To explore, clean, analyse, and visualise the data through a dashboard.


### Data Cleaning and Transformation
After importing the dataset into Excel and converting the file into an Excel Worksheet, I converted the data into an Excel Table to manage and analyse the data easier. The dataset already contained a `year` column, but I considered best to use an Excel formula to extract the year from the `order_date` colum. First, I changed the format of `order_date` column into a Date, then I used an Excel formula to extract the year, formatted the year as a number with no decimal, and named the colum `order_year`.

In most business each customer has a customer ID assigned, but in this dataset that variable is missing. However, using the `customer_name` variable, it is possible to identify unique customers using a conditional Excel formula to count each time a new customer is added to the database. 

Another variable to count was the unique number of orders. In this case, the database has a `order_id` column, so I used a similar conditional formula to count the unique number of `order_id`s. 










### Data Analysis
Using a Pivot Table, first I calculated the Total Sales, the Total Number of Unique Orders, and the Total Number of Unique Customers. Then, I created a calculated field to estimate the Average Sales per Order, leaving the result with two decimals.
Then, I added a slicer in the `order_year` to filter the Pivot Table and the final dashboard.
  
I used two pivot tables to calculate the total sales by quarter, and year using the `order_year` variable.  

### Excel formulas
| Formula | Description | 
| :--- | :--- |
| `=YEAR([@[order_date]])` | Extract the Year from the order date |
| `=IF(COUNTIF(F$2:F2,F2)=1,1,0) ` | Count each unique customer. If the user is unique, it assigns a 1 value, if the customer is repeated, it assigs a 0.  |
| `=IF(COUNTIF(A$2:A2,A2)=1,1,0) ` | Count each unique order_id. If an order_id is unique, it assigns a 1 value, if the order_id is repeated, it assigs a 0. |
| `Avg Sales per Order = sales/ order_unique ` | A calculated field to estimate the average sales per order  |
| ` ` |  |
| ` ` |  |
| ` ` |  |
| ` ` |  |


#### Data Visualisation




## Insights

## Action plan


## Excel Report
To review the analysis in detail, you can download the Excel Report [here]().
