# 📊 Excel Data Exploration & Analysis

## 📌 Project Overview

This project demonstrates my foundational **Data Analysis and Microsoft Excel skills** through the exploration and analysis of a product dataset.

The objective of this project was to perform basic data exploration, apply Excel formulas and functions, categorize products based on price, perform conditional calculations, and extract structured information from Product IDs.

This project is part of my journey toward becoming a **Data Analyst** and serves as one of the projects in my data analytics portfolio.

---

## 📂 Dataset

The dataset contains information about different products with the following attributes:

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Number of units                    |
| Category     | Product category                   |
| Price        | Price of the product               |

---

## 🎯 Project Objectives

The main objectives of this project were to:

* Explore and summarize product data
* Calculate key statistical measures
* Identify minimum and maximum product prices
* Categorize products based on price
* Perform conditional aggregation
* Extract information from Product IDs
* Practice Excel formulas and functions
* Build a structured analytical worksheet

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Formulas & Functions
* Data Exploration
* Data Summarization
* Logical Analysis
* Conditional Aggregation
* Text Manipulation

---

# 🔍 Analysis Performed

## 1. Basic Data Exploration

The following Excel functions were used to summarize the dataset:

### SUM

Calculated the **total price of all products**.

```excel
=SUM(F2:F[n])
```

### COUNT

Calculated the **number of products** in the dataset.

```excel
=COUNT(F2:F[n])
```

### AVERAGE

Calculated the **average product price**.

```excel
=AVERAGE(F2:F[n])
```

These calculations provide an initial overview of the numerical characteristics of the dataset.

---

## 2. Minimum & Maximum Price Analysis

The `MIN` and `MAX` functions were used to identify the lowest and highest product prices.

### Minimum Price

```excel
=MIN(F2:F[n])
```

### Maximum Price

```excel
=MAX(F2:F[n])
```

This provides a basic understanding of the price range within the product dataset.

---

## 3. Price Categorization Using IF

A new column named **Price Range** was created to categorize products according to their price.

### Business Rule

* Price **≥ $500** → `High Price`
* Price **< $500** → `Standard Price`

### Formula

```excel
=IF(F2>=500,"High Price","Standard Price")
```

The formula was then applied to all product records.

This demonstrates the use of **logical conditions in Excel** to transform numerical data into meaningful categories.

---

## 4. Conditional Analysis

### SUMIF – Electronics Products

The `SUMIF` function was used to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(E2:E[n],"Electronics",F2:F[n])
```

This demonstrates how Excel can be used to perform calculations based on specific conditions.

### COUNTIF – Products Below $100

The `COUNTIF` function was used to determine the number of products with a price below $100.

```excel
=COUNTIF(F2:F[n],"<100")
```

This provides an example of conditional counting using Excel.

---

# 🔤 5. Text Manipulation Using LEFT, RIGHT & MID

The Product ID was analyzed using Excel text functions.

Three new columns were created:

### Day

The first two characters of the Product ID were extracted using `LEFT`.

```excel
=LEFT(A2,2)
```

### Country Code

The last two characters of the Product ID were extracted using `RIGHT`.

```excel
=RIGHT(A2,2)
```

### Month

Characters 4 to 6 of the Product ID were extracted using `MID`.

```excel
=MID(A2,4,3)
```

These functions demonstrate basic **text transformation and data extraction**, which are useful when working with structured identifiers and datasets.

---

# 📈 Skills Demonstrated

This project demonstrates the following Data Analyst skills:

### Data Exploration

* SUM
* COUNT
* AVERAGE
* MIN
* MAX

### Logical Analysis

* IF statements
* Rule-based categorization

### Conditional Aggregation

* SUMIF
* COUNTIF

### Text Processing

* LEFT
* RIGHT
* MID

### Spreadsheet Skills

* Formula application
* Cell referencing
* Creating calculated columns
* Organizing analytical results
* Basic Excel data analysis

---

# 📁 Project Structure

```text
Excel-Data-Exploration/
│
├── 📊 Excel Assignment 1 - Data Exploration.xlsx
│
├── 📄 Documentation/
│   └── Excel Data Exploration Screenshots.pdf
│
└── README.md
```

---

# 📸 Documentation

The project documentation includes screenshots demonstrating:

* Basic calculation results
* Minimum and maximum price calculations
* Price Range classification
* SUMIF calculation
* COUNTIF calculation
* LEFT function
* RIGHT function
* MID function
* Applied formulas visible in the Excel formula bar

The screenshots provide evidence of the formulas and their application within the workbook.

---

# 💡 Key Learning Outcomes

Through this project, I strengthened my understanding of:

* Working with structured datasets in Excel
* Using formulas for data summarization
* Applying logical conditions to datasets
* Performing conditional calculations
* Extracting information from text fields
* Using relative cell references
* Organizing analytical results
* Documenting data analysis work

---

# 🚀 Portfolio Context

This project represents one of my **foundational Data Analytics projects** as I build my professional portfolio.

I am currently developing my skills across:

* Microsoft Excel
* SQL
* Python
* Data Cleaning
* Data Visualization
* Exploratory Data Analysis
* Power BI

My goal is to progress from foundational spreadsheet analysis toward more advanced **data analytics and business intelligence projects**.

---

# 👨‍💻 Author

**Gokul Sree**

Aspiring Data Analyst

This repository is part of my Data Analytics learning and portfolio journey.

---

## ⭐ Future Improvements

As I continue developing my analytical skills, future versions of this project could include:

* Excel Pivot Tables
* Pivot Charts
* Interactive dashboards
* Additional data quality checks
* More advanced Excel functions
* SQL-based analysis of the same dataset
* Python-based exploratory data analysis
* Power BI visualization

---

## 📌 Conclusion

This project demonstrates how Microsoft Excel can be used to perform fundamental data exploration, summarization, conditional analysis, logical categorization, and text manipulation.

It provides a foundation for progressing toward more advanced data analysis projects using **SQL, Python, Excel, and Power BI**.
-1
