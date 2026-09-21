# 📊 Excel Data Exploration & Analysis

## 📌 Project Overview

This project is my first step toward building a **Data Analytics portfolio**. The objective of this assignment is to perform basic **data exploration and analysis using Microsoft Excel**.

The project uses a product dataset containing information such as **Product ID, Product Name, Brand Name, Price, Quantity, and Category**. I used Excel formulas and functions to summarize the data, apply conditional logic, perform conditional calculations, and extract useful information from text fields.

This project demonstrates foundational Excel skills that are important for **Data Analysis and Business Intelligence**.

---

## 🎯 Objectives

The main objectives of this project are:

* Explore and summarize a product dataset.
* Calculate total, count, average, minimum, and maximum values.
* Categorize products based on their price.
* Perform conditional calculations using `SUMIF` and `COUNTIF`.
* Extract information from Product IDs using text functions.
* Practice Excel formulas commonly used in data analysis.
* Build a foundation for future data analytics projects.

---

## 📂 Dataset

**Dataset:** Product Dataset

### Dataset Attributes

| Column       | Description                                               |
| ------------ | --------------------------------------------------------- |
| Product ID   | Unique identifier containing day, month, and country code |
| Product Name | Name of the product                                       |
| Brand Name   | Product brand                                             |
| Price ($)    | Price of the product                                      |
| Quantity     | Quantity of products                                      |
| Category     | Product category                                          |

The dataset contains **34 product records**.

---

# 🔎 Tasks Performed

## 1. Basic Data Exploration

### Total Price

Excel formula:

```excel
=SUM(D2:D35)
```

**Result: $10,100**

### Number of Products

Excel formula:

```excel
=COUNTA(B2:B35)
```

**Result: 34 products**

### Average Price

Excel formula:

```excel
=AVERAGE(D2:D35)
```

**Result: $297.06**

---

## 2. Minimum and Maximum Price

### Minimum Price

```excel
=MIN(D2:D35)
```

**Minimum Price: $30**

### Maximum Price

```excel
=MAX(D2:D35)
```

**Maximum Price: $1,000**

---

## 3. Price Range Using IF

A new column named **Price Range** was created using the `IF` function.

### Business Rule

* Price **≥ $500** → `High Price`
* Price **< $500** → `Standard Price`

### Formula

```excel
=IF(D2>=500,"High Price","Standard Price")
```

This formula was filled down for all product records.

---

## 4. Conditional Analysis

### Total Price of Electronics Products

The `SUMIF` function was used to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

**Result: $8,050**

### Number of Products Priced Below $100

The `COUNTIF` function was used to count products with a price less than $100.

```excel
=COUNTIF(D2:D35,"<100")
```

**Result: 11 products**

---

# 🔤 5. Text Functions

The Product ID follows a structure similar to:

```text
28-JAN-US
```

It contains:

* `28` → Day
* `JAN` → Month
* `US` → Country Code

Three new columns were created using Excel text functions.

### Day – LEFT

The first 2 characters were extracted using `LEFT`.

```excel
=LEFT(A2,2)
```

Example:

```text
28-JAN-US → 28
```

---

### Country Code – RIGHT

The last 2 characters were extracted using `RIGHT`.

```excel
=RIGHT(A2,2)
```

Example:

```text
28-JAN-US → US
```

---

### Month – MID

Characters 4 to 6 were extracted using `MID`.

```excel
=MID(A2,4,3)
```

Example:

```text
28-JAN-US → JAN
```

---

# 📈 Project Results

| Analysis                |  Result |
| ----------------------- | ------: |
| Total Products          |      34 |
| Total Price             | $10,100 |
| Average Price           | $297.06 |
| Minimum Price           |     $30 |
| Maximum Price           |  $1,000 |
| Electronics Total Price |  $8,050 |
| Products Below $100     |      11 |

---

# 🛠️ Excel Functions Used

The following Excel functions were used in this project:

```text
SUM
COUNTA
AVERAGE
MIN
MAX
IF
SUMIF
COUNTIF
LEFT
RIGHT
MID
```

---

# 💡 Skills Demonstrated

Through this assignment, I practiced:

* Data Exploration
* Data Summarization
* Excel Formulas
* Conditional Logic
* Conditional Aggregation
* Text Manipulation
* Data Categorization
* Basic Data Analysis
* Spreadsheet Data Cleaning and Preparation

---

# 📁 Project Files

The repository contains:

```text
Excel-Data-Exploration/
│
├── README.md
│
└── Excel Assignment 1-Data Exploration working.xlsx
```

---

# 🚀 Learning Outcome

This project helped me understand how Excel can be used to perform fundamental data analysis tasks.

I learned how to use formulas to **summarize numerical data, apply business rules, perform conditional calculations, and extract structured information from text fields**.

This assignment is part of my journey toward building a **Data Analytics portfolio** and developing practical skills in **Excel, data analysis, and business intelligence**.

---

## 🔜 Future Projects

I plan to continue expanding my portfolio with projects involving:

* Advanced Excel
* Power BI
* SQL
* Python
* Data Cleaning
* Data Visualization
* Exploratory Data Analysis
* Business Intelligence Dashboards

---

## 👤 About Me

I am currently building my skills in **Data Analytics** through hands-on projects and practical assignments.

This repository documents my learning journey and demonstrates my progress in applying data analysis techniques to real-world style datasets.
