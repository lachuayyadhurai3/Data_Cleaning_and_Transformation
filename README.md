# 📊 Data Cleaning and Transformation using Excel

A simple project that cleans and transforms a small retail product dataset (34 rows) using only Excel.

## 📁 Dataset

| Column | Description |
|---|---|
| Product ID | Code in the format `DD-MON-CC` (day, month, country code) |
| Product Name | Name of the product |
| Brand Name | Brand of the product |
| Price ($) | Price in US dollars |
| Quantity | Units in stock |
| Category | Product category |

## Data Quality Issues Found

| Issue | Example | Fix |
|---|---|---|
| Missing Price | Sony Headphones, Coleman Camping Tent, Ray-Ban Sunglasses | Fill with the median |
| Missing Category | Backpack, Sneakers, Coffee Maker, Fitness Tracker | Fill based on the product type |
| Truncated text | `Electroni` instead of `Electronics` | Find & Replace |
| Inconsistent case | `laptop`, `Laptop`, `headphones`, `smartphone` | Standardise with `PROPER` |
| Duplicate rows | Laptop Bag (Samsonite), Headphones (Bose), Laptop (HP) | Remove duplicates |
| Combined values | Product ID holds both a date and a country | Split into two columns |

## Steps

### 1. Make a backup
Copy the raw sheet to a new sheet named `Raw_Dataset`, do all the cleaning on a separate sheet named `Working_Dataset` and cleaned dataset sheet named `Cleaned_Dataset`.

### 2. Remove duplicates
`Data > Remove Duplicates`, select all columns, and click OK.
This removes 3 duplicate rows and leaves 31 records.

### 3. Fix text formatting
Use these formulas in a helper column, then paste the results back as values:

    =PROPER(TRIM(B2))

This makes `laptop` and `Laptop` the same value and removes extra spaces.

### 4. Fix the truncated category
Press `Ctrl + H` and replace `Electroni` with `Electronics`.
Tick **Match entire cell contents**, otherwise the existing `Electronics` values become `Electronicscs`.

### 5. Handle missing Category
Fill the blanks according to the product type:

| Product | Category |
|---|---|
| Backpack | Outdoor |
| Sneakers | Fashion |
| Coffee Maker | Kitchen |
| Fitness Tracker | Electronics |

### 6. Handle missing Price
Prices are right-skewed (a few expensive items such as laptops), so use the **median**, not the mean.
The median is about 130, while the mean is about 298 and would be pulled up by the expensive items.

    =MEDIAN(D2:D32)

Fill each blank Price cell with this value.
For a more accurate result, use the median of the same Category instead.

### 7. Split Product ID (transformation)

| New column | Formula |
|---|---|
| Day_Month | `=LEFT(A2,6)` |
| Country | `=RIGHT(A2,2)` |

### 8. Final checks
- No blank cells (`Ctrl + G > Special > Blanks`)
- No duplicates
- Consistent text case and category names
- Correct data types (numbers right-aligned)

## 🛠️ Excel Features Used
- Remove Duplicates
- Find & Replace
- `TRIM`, `PROPER`, `LEFT`, `RIGHT`, `MEDIAN`
- Go To Special (to find blanks)
- Filters and Pivot Tables (for checking the results)

## 📌 Result
A clean dataset with no duplicates, no missing values, standardised text, and three new transformed columns (`Country Code`, `Manufacturing Date`, `Product Brand`).

## 🗂️ Files
- `Raw_DataSet.xlsx`: original dataset
- `Working_Dataset.xlsx`: Worked Document
- `Cleaned_Dataset.xlsx`: cleaned and transformed dataset
- `README.md`: project documentation
