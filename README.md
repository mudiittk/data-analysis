# Global Sales Insights (2012-2015) - Power BI Dashboard

## Overview
This project provides an in-depth analysis of global sales data from 2012 to 2015 using Power BI. The dashboard offers key insights into sales performance, profitability, shipping trends, and product performance. It helps identify high and low-performing segments, visualize sales trends over time, and analyze key performance metrics to support data-driven decision-making.

## Files in the Repository
- **Amazon-Dataset(AutoRecovered).xlsx**: This Excel file contains the raw sales data used for analysis. Although the data was already clean, various operations were performed in Excel before importing it into Power BI to enhance usability. The original dataset is from here: https://github.com/DishaK06/Amazon-Global-Sales
- **Global Sales Insight (2012-2015).pbix**: This is the Power BI report file that contains all the visualizations, transformations, DAX functions and analytical insights.
- **Theme.json**: This is the theme which i used in my dashboard which i downloaded from here : https://community.fabric.microsoft.com/t5/Themes-Gallery/Sales-Report/td-p/4398953
## Data Preparation
While the dataset was already clean, the following operations were performed in Excel before importing it into Power BI:
- **Pivot Tables and Charts**: Used to summarize and analyze key sales trends before importing into Power BI.
- **Creating a New Column for Delivery Time Calculation**: A new column was created by calculating the difference between the ship date and order date to determine the number of days taken to deliver a product.
- **Shipping Mode Analysis**: Used ship mode categories to calculate the average number of days taken to deliver products under each category.
  - If the difference between ship date and order date was larger than the average, it indicated late delivery.
  - If the difference matched the average, it indicated on-time delivery.
  - If the difference was smaller than the average, it indicated early delivery.
- **Delivery Performance Analysis**: Using the above calculations, the total percentage of orders that were delivered on time, before time, and late was determined.
- **Growth Rate Calculation**: Pivot tables were used to calculate the growth rate for sales, profit, and quantity over the years.

## Key Dashboard Features
### 1. **Sales and Profit Overview**
- **Total Sales**: Displays overall revenue generated.
- **Average Order Value (AOV)**: Measures the average transaction value.
- **Profit & Profit Margin**: Shows overall profitability trends.

### 2. **Sales Analysis by Time**
- **Monthly Sales vs. Costs**: A bar chart comparing monthly revenue and expenses.
- **Profit Over Time**: A trend line showing profitability growth.
- **Sales Over Time**: A line chart depicting year-over-year sales growth.

### 3. **Geographical Insights**
- **Sales by Country**: A horizontal bar chart ranking countries by revenue contribution.
- **Regional Sales Map**: A geographical map visualization of sales distribution.

### 4. **Product and Category Performance**
- **Most & Least Profitable Products**: Identifies high and low-performing products.
- **Profit by Segment**: Breaks down profit by Home Office, Corporate, and Consumer segments.
- **Profit by Category**: Shows profitability for Furniture, Office Supplies, and Technology.
- **Discount by Sub-Category**: Highlights discount rates across different product categories.

### 5. **Shipping Mode and Performance**
- **Profit by Ship Mode**: Compares profitability across First Class, Second Class, and Standard shipping.
- **Sales vs. Profit Scatter Plot**: Identifies high and low-performing sales regions and products.

## DAX Functions Used
Below are some of the key DAX functions used in the Power BI dashboard:

### **AOV Color (Conditional formatting for highlighting months with highest/lowest AOV)**
```DAX
AOV Color = 
VAR currMonthAOV = CALCULATE(
    [Average Order Value],
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)
VAR maxMonthAOV = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), [Average Order Value]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)
VAR minMonthAOV = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), [Average Order Value]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)
RETURN  
IF(currMonthAOV = maxMonthAOV, "#418f87", 
    IF(currMonthAOV = minMonthAOV, "Red", "DCE5ED")
)
```
### **AOV Label** (Conditional Formatting to display only the highest/lowest month values)
```DAX
AOV Label = 
VAR currMonthAOV = CALCULATE(
    [Average Order Value],
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)

VAR maxMonthAOV = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), [Average Order Value]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

VAR minMonthAOV = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), [Average Order Value]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

RETURN  
IF(currMonthAOV = maxMonthAOV || currMonthAOV = minMonthAOV, currMonthAOV, BLANK())
```

### **Average Order Value Calculation**
```DAX
Average Order Value = SUM(Orders[Sales])/SUM(Orders[Quantity])
```

### **Cost Calculation**
```DAX
Cost = SUM(Orders[Sales])-SUM(Orders[Profit])
```

### **Max & Min Sales Calculation**
```DAX
Max_sales = MAX('Orders'[Sales])
Min_sales = MIN('Orders'[Sales])
```

### **Profit Margin Calculation**
```DAX
Profit Margin = SUM(Orders[Profit])/SUM(Orders[Sales])
```

### **Previous Year Sales Calculation**
```DAX
Previous year sales = CALCULATE([Total Sales], SAMEPERIODLASTYEAR('Orders'[Order Date]))
```

### **Total Sales Calculation**
```DAX
Total Sales = SUM(Orders[Sales])
```

### **Sales Color (Conditional Formatting for highlighting months with the highest/lowest total sales)**
```DAX
Sales Color = 
VAR currMonthSales = CALCULATE(SUM(Orders[Sales]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)
VAR maxMonthSales = CALCULATE(MAXX(VALUES(Orders[Order Date].[Month]), SUM(Orders[Sales])), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)
VAR minMonthSales = CALCULATE(MINX(VALUES(Orders[Order Date].[Month]), SUM(Orders[Sales])), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)
RETURN  
IF(currMonthSales = maxMonthSales, "#418f87",  
    IF(currMonthSales = minMonthSales, "Red", "DCE5ED")
)
```
### **Sales Label(Conditional Formatting to display only the highest/lowest month values)**
```DAX
Total Sales Label = 
VAR currMonthSales = CALCULATE(
    Orders[Total Sales],
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)

VAR maxMonthSales = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), Orders[Total Sales]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

VAR minMonthSales = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), Orders[Total Sales]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

RETURN  
IF(currMonthSales = maxMonthSales || currMonthSales = minMonthSales, currMonthSales, BLANK())

```
### **Profit Color (Conditional Formatting for highlighting months with the highest/lowest total profit)**
```DAX
Profit Color = 
VAR currMonthProfit = CALCULATE(
    SUM(Orders[Profit]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)

VAR maxMonthProfit = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), CALCULATE(SUM(Orders[Profit]))), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

VAR minMonthProfit = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), CALCULATE(SUM(Orders[Profit]))), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

RETURN  
IF(currMonthProfit = maxMonthProfit, "#418f87",  
    IF(currMonthProfit= minMonthProfit, "Red", "DCE5ED")
)
```
### **Profit Label (Conditional Formatting to display only the highest/lowest month values)**
```DAX
Profit Label = 
VAR currMonthProfit = CALCULATE(
    SUM(Orders[Profit]),
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)

VAR maxMonthProfit = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), 
        CALCULATE(SUM(Orders[Profit]))), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

VAR minMonthProfit = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), 
        CALCULATE(SUM(Orders[Profit]))), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

RETURN  
IF(currMonthProfit = maxMonthProfit || currMonthProfit = minMonthProfit, currMonthProfit, BLANK())
```
### **Profit Margin Color(Conditional Formatting for highlighting months with the highest/lowest profit margin)**
```DAX
PM color = 
VAR currMonthPM = CALCULATE(
    [Profit Margin],
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)

VAR maxMonthPM = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), [Profit Margin]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

VAR minMonthPM = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), [Profit Margin]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)
RETURN  
IF(currMonthPM = maxMonthPM, "#418f87",  
    IF(currMonthPM= minMonthPM, "Red", "DCE5ED")  
)
```
### **Prfoit Margin Label (Conditional Formatting to display only the highest/lowest month values)**
```DAX 
PM Label = 
VAR currMonthPML = CALCULATE(
    [Profit Margin],
    ALLEXCEPT(Orders, Orders[Order Date].[Year], Orders[Order Date].[Month])
)

VAR maxMonthPML = CALCULATE(
    MAXX(VALUES(Orders[Order Date].[Month]), [Profit Margin]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

VAR minMonthPML = CALCULATE(
    MINX(VALUES(Orders[Order Date].[Month]), [Profit Margin]), 
    ALLEXCEPT(Orders, Orders[Order Date].[Year])
)

RETURN  
IF(currMonthPML = maxMonthPML || currMonthPML = minMonthPML, currMonthPML, BLANK())

```

## Technologies Used
- **Power BI**: Used for data visualization and reporting.
- **Excel**: Used for initial data formatting, analysis and structuring.

## How to Use This Dashboard
1. **Download the PBIX File**
   - Open the `Global Sales Insight (2012-2015).pbix` file in Power BI Desktop.
2. **Interact with Filters and Slicers**
   - Use year filters (2012-2015) to analyze trends for a specific period.
   - Hover over visuals to explore detailed insights.
3. **Explore Key Insights**
   - Identify top-performing countries, categories, and shipping modes.
   - Analyze discount strategies and their impact on profitability.
