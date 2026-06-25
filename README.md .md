# 📊 Excel Statistical Analysis — Sales Data

**Tool Used:** Microsoft Excel  
**Analysis Types:** Correlation Analysis · Multiple Linear Regression  
**Dataset:** 100 e-commerce sales orders across the Middle East & North Africa region

---

## 📌 Project Overview

This project applies statistical analysis techniques in Excel to a real-world e-commerce sales dataset. The goal was to uncover relationships between business variables and build a predictive model for total sales — combining data cleaning, correlation analysis, and multiple regression, all within Excel.

The dataset contains 100 orders placed across Jordan, Saudi Arabia, UAE, and Egypt, with fields covering customer behavior, product details, payment methods, order status, and delivery performance.

---

## 📂 File Structure

| Sheet Name            | Description                                                                 |
|-----------------------|-----------------------------------------------------------------------------|
| `Sales_Data`          | Raw dataset (100 orders) with applied text/data cleaning formulas           |
| `Correlation`         | Correlation analysis: Customer Rating vs. Total Sales                       |
| `Regression`          | Data preparation and cleaning for regression inputs                         |
| `Regression clean data` | Multiple regression model output and ANOVA summary table                  |

---

## 📋 Dataset Overview

The `Sales_Data` sheet contains 100 rows and 33+ columns, including:

| Field | Description |
|-------|-------------|
| `Order_ID` | Unique order identifier |
| `Customer_Name / Email / Phone` | Customer contact details (with intentional data quality issues) |
| `Country / City / Sales_Region` | Geographic data (Jordan, Saudi Arabia, UAE, Egypt) |
| `Product_Name / Category / Brand` | Product details (Electronics, Furniture, Home Appliances, Accessories) |
| `Order_Date / Delivery_Date` | Date fields for delivery time calculation |
| `Quantity / Unit_Price / Discount_Percentage` | Order financials |
| `Shipping_Fee / Total_Sales` | Cost and revenue fields |
| `Customer_Rating` | Satisfaction score (1–5 scale) |
| `Order_Status` | Delivered, Pending, or Returned |
| `Payment_Method / Sales_Channel` | Cash, Credit Card, PayPal, Bank Transfer; Store, Online, Mobile App |
| `Customer_Feedback` | Free-text feedback field (intentionally messy) |
| `Verification_Status / Error_Check` | Data integrity flags |

### 🔧 Data Quality Issues Applied (for cleaning practice)
The dataset was intentionally built with real-world data problems, including:
- Extra spaces in emails and feedback fields
- Invalid phone numbers (`07A123456`, `N/A`)
- Missing values in Quantity and Customer_Rating
- Inconsistent email formats (missing `@`)
- `#N/A` and `#VALUE!` formula errors in derived columns

---

## 🔍 Analyses Performed

### 1. Data Cleaning (Sales_Data Sheet)
A set of Excel formulas was applied across added columns to audit and clean the raw data:

| Formula Task | Excel Function Used |
|---|---|
| Extract last 4 characters of Tracking Number | `RIGHT()` |
| Count characters in Customer Feedback | `LEN()` |
| Remove extra spaces from Feedback | `TRIM()` |
| Combine Order ID, Product, and Country into a label | `CONCAT()` / `&` |
| Find position of `@` in email | `FIND()` |
| Convert text-based quantities to numbers | `IFERROR()` + `VALUE()` |
| Classify delivery as Late or On Time (>5 days) | `IF()` + date arithmetic |
| Classify Total_Sales into tiers (Bronze/Silver/Gold/Platinum) | Nested `IFS()` |
| Convert numeric Priority_Level to text labels | `IFS()` |
| Count Delivered orders | `COUNTIF()` |
| Count GCC orders paid by Credit Card | `COUNTIFS()` |
| Count numeric ratings | `COUNT()` |
| Count non-empty Notes | `COUNTA()` |
| Count blank Feedback entries | `COUNTBLANK()` |
| Validate Quantity as numeric | `ISNUMBER()` |
| Check Feedback data type | `ISTEXT()` |
| Detect blank ratings | `ISBLANK()` |

---

### 2. Correlation Analysis (Correlation Sheet)

**Question:** Is there a relationship between customer satisfaction and sales performance?

**Variables:**
- Independent: `Customer_Rating` (1–5 scale)
- Dependent: `Total_Sales` (monetary value)

**Method:** Pearson Correlation Coefficient using Excel's `CORREL()` function after cleaning missing values.

**Result:**

| Metric | Value |
|--------|-------|
| Correlation Coefficient (r) | **-0.155** |
| Relationship Strength | Very weak negative |

**Interpretation:** There is virtually no meaningful linear relationship between customer ratings and total sales in this dataset. A high customer rating does not predict a high (or low) sales value, suggesting that satisfaction and revenue are driven by independent factors in this context.

---

### 3. Multiple Linear Regression (Regression & Regression Clean Data Sheets)

**Question:** Can Quantity, Discount Percentage, Shipping Fee, and Customer Rating predict Total Sales?

**Dependent Variable:** `Total_Sales`  
**Independent Variables:** `Quantity`, `Discount_Percentage`, `Shipping_Fee`, `Customer_Rating`

**Data Cleaning Steps:**
- Removed rows with `N/A` in Quantity (non-numeric entries)
- Replaced blank Customer_Rating entries with `0`
- Verified all columns are numeric before running the Excel Data Analysis ToolPak regression

**Final sample size after cleaning:** 51 observations

**Regression Output Summary:**

| Metric | Value |
|--------|-------|
| Multiple R | 0.500 |
| R² | 0.250 |
| Adjusted R² | 0.185 |
| Standard Error | 2,396.98 |
| Model p-value (Significance F) | **0.009** ✅ Significant |

**Coefficient Table:**

| Predictor | Coefficient | p-value | Significant? |
|-----------|-------------|---------|--------------|
| Intercept | 2,066.90 | 0.140 | No |
| Quantity | **+369.13** | **0.003** | ✅ Yes |
| Discount_Percentage | -78.73 | 0.087 | No |
| Shipping_Fee | -11.48 | 0.862 | No |
| Customer_Rating | -232.27 | 0.306 | No |

**Interpretation:**
- The model is statistically significant overall (p = 0.009) and explains **25% of the variance** in Total Sales.
- **Quantity is the only significant predictor**: each additional unit ordered is associated with an average increase of **~369 JOD** in total sales, holding other variables constant.
- Discount percentage, shipping fee, and customer rating showed no statistically significant effect on total sales in this dataset.
- The moderate R² (0.25) suggests other variables not captured in this dataset (e.g., product category, seasonality, customer segment) likely explain the remaining variance.

---

## 💡 Key Insights

- Customer satisfaction scores **do not reliably predict sales revenue** in this dataset (r = -0.155).
- **Order quantity is the strongest driver of total sales**, which aligns with basic revenue logic (Revenue = Price × Quantity).
- Discounts and shipping fees, while commonly assumed to influence purchase behavior, showed no statistically significant impact on total sales here — possibly due to the limited sample size (n = 51 after cleaning).
- The dataset contains rich data quality issues that make it a useful practice ground for real-world data cleaning scenarios.

---

## 🛠 Tools & Skills Used

- Microsoft Excel (Data Analysis ToolPak — Regression)
- `CORREL()`, `COUNTIF()`, `COUNTIFS()`, `TRIM()`, `LEN()`, `RIGHT()`, `FIND()`, `IFERROR()`, `VALUE()`, `IF()`, `IFS()`, `ISNUMBER()`, `ISTEXT()`, `ISBLANK()`
- Pearson Correlation Analysis
- Multiple Linear Regression (OLS)
- Data cleaning and validation techniques

---

## 📁 How to Use

1. Download the `.xlsx` file.
2. Open in **Microsoft Excel** (2016 or later).
3. To re-run the regression, go to **Data → Data Analysis → Regression** and select the cleaned columns from the `Regression clean data` sheet.
4. The correlation matrix is in the `Correlation` sheet, computed using `=CORREL()`.
5. All cleaning formulas are visible in the added columns on the right side of `Sales_Data`.

---

## 📬 Contact

**Fakhri Abuhantash**  
BSc Business Intelligence & Data Analytics  
German Jordanian University  
