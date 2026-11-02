# Excel Data Manipulation

## 1. Overview and Fundamentals

### 1.1 What Is Data Manipulation?

Data Manipulation is the process of modifying, organizing, restructuring and transforming data to make it useful for analysis and reporting.

It involves operations such as:
- Selecting specific rows and columns.
- Sorting records.
- Filtering records.
- Adding or removing columns.
- Combining datasets.
- Transforming values.
- Reshaping data.
- Extracting required information.

Example:

An employee dataset may contain 1,000 records. A business analyst may need to extract employees from the IT department, sort them by salary and calculate department-wise totals.

These operations are examples of data manipulation.

### 1.2 Why Is Data Manipulation Important?

Data manipulation helps to:

- Organize large datasets.
- Extract relevant information.
- Prepare data for analysis.
- Combine information from different sources.
- Transform raw data into useful formats.
- Support business decision-making.
- Prepare reports and dashboards.

### 1.3 Data Manipulation vs. Data Cleaning

| Data Manipulation | Data Cleaning |
|---|---|
| Changes data structure or arrangement | Improves data quality |
| Selects, filters and transforms records | Handles errors and inconsistencies |
| Combines and reshapes datasets | Handles duplicates and missing values |
| Prepares data for analysis | Makes data reliable for analysis |

Both processes are commonly used together.

### 1.4 Common Data Manipulation Operations

1. Selecting rows and columns.
2. Sorting records.
3. Filtering records.
4. Adding and deleting columns.
5. Extracting values.
6. Replacing values.
7. Combining datasets.
8. Splitting columns.
9. Grouping records.
10. Reshaping data.

---

## 2. Core Concepts, Methods and Techniques

### 2.1 Selecting Rows and Columns

Selection means retrieving specific parts of a dataset.

Example:

| Employee | Department | Salary |
|---|---|---:|
| Arun | IT | 30000 |
| Priya | HR | 25000 |
| Kumar | IT | 35000 |

Possible operations:
- Select only Employee and Salary columns.
- Select only IT employees.
- Select employees earning more than 30000.

Excel methods:
- Select worksheet ranges.
- Use filters.
- Use formulas such as FILTER.
- Use Power Query.

### 2.2 Sorting Data

Sorting arranges records in a particular order.

Common sorting methods:
- A to Z.
- Z to A.
- Smallest to Largest.
- Largest to Smallest.
- Oldest to Newest.
- Newest to Oldest.
- Custom sorting.

Example:

Sort employee salaries from highest to lowest.

| Employee | Salary |
|---|---:|
| Kumar | 35000 |
| Arun | 30000 |
| Priya | 25000 |

### 2.3 Filtering Data

Filtering displays only records that satisfy specified conditions.

Example:

Display only employees from the IT department.

Excel provides:
- AutoFilter.
- Text Filters.
- Number Filters.
- Date Filters.
- Advanced Filter.

The original dataset remains available while filtered records are displayed.

### 2.4 Adding and Removing Columns

Adding columns helps create new information.

Example:

Original columns:
- Product
- Quantity
- Price

New column:
- Total Amount

Formula:

```excel
=B2*C2
```

Removing columns eliminates unnecessary information.

Before deleting columns, confirm that they are not required by formulas, charts or reports.

### 2.5 Extracting Data

Data extraction retrieves specific information from existing values.

Examples:
- Extract the first name from a full name.
- Extract the year from a date.
- Extract the domain from an email address.
- Extract product codes from a description.

Useful Excel functions:
- LEFT
- RIGHT
- MID
- FIND
- SEARCH
- TEXTBEFORE
- TEXTAFTER
- TEXTSPLIT

### 2.6 Replacing Values

Find and Replace changes selected values in a dataset.

Example:

Replace:

```text
IT Department → Information Technology
```

Excel shortcut:

```text
Ctrl + H
```

Use Find and Replace carefully because replacing a value across an entire worksheet may affect unrelated information.

### 2.7 Combining Datasets

Combining datasets means bringing information from multiple sources together.

Common methods:

**Vertical Combination**

Append rows from multiple datasets with similar columns.

Example:
- January Sales
- February Sales
- March Sales

Combine all monthly records into one sales dataset.

**Horizontal Combination**

Combine columns using a shared identifier.

Example:

Employee details + Employee salary records.

Common Excel methods:
- Power Query Append.
- Power Query Merge.
- XLOOKUP.
- VLOOKUP.
- INDEX and MATCH.

### 2.8 Grouping Data

Grouping organizes records into categories.

Examples:
- Sales by region.
- Employees by department.
- Revenue by month.
- Products by category.

Useful Excel methods:
- PivotTables.
- SUMIF.
- SUMIFS.
- COUNTIF.
- COUNTIFS.

### 2.9 Data Transformation

Data transformation converts information into a different structure or representation.

Examples:
- Convert text to uppercase.
- Convert dates into month and year.
- Convert wide data into long data.
- Create calculated columns.
- Standardize category names.
- Convert units.

Example:

```excel
=UPPER(A2)
```

This converts text into uppercase.

### 2.10 Power Query for Data Manipulation

Power Query is an Excel data transformation and preparation tool.

Common operations:
- Remove columns.
- Filter rows.
- Change data types.
- Split columns.
- Merge queries.
- Append queries.
- Group rows.
- Pivot columns.
- Unpivot columns.

Power Query is especially useful for repeatable data preparation workflows.

---

## 3. Step-by-Step Implementation with Excel Examples

### 3.1 Example: Filter IT Employees

Dataset:

| Employee | Department | Salary |
|---|---|---:|
| Arun | IT | 30000 |
| Priya | HR | 25000 |
| Kumar | IT | 35000 |
| Divya | Finance | 28000 |

Requirement: Display only IT employees.

Steps:
1. Select the dataset.
2. Open the Data tab.
3. Click Filter.
4. Open the Department dropdown.
5. Select IT.
6. Click OK.

Expected result:

| Employee | Department | Salary |
|---|---|---:|
| Arun | IT | 30000 |
| Kumar | IT | 35000 |

### 3.2 Example: Sort Sales by Highest Value

Dataset:

| Product | Sales |
|---|---:|
| Laptop | 85000 |
| Mobile | 45000 |
| Monitor | 120000 |

Steps:
1. Select the dataset.
2. Open Data.
3. Select Sort.
4. Choose Sales as the sorting column.
5. Select Largest to Smallest.
6. Click OK.

Expected order:
1. Monitor — 120000
2. Laptop — 85000
3. Mobile — 45000

### 3.3 Example: Create a Total Amount Column

Dataset:

| Product | Quantity | Price |
|---|---:|---:|
| Laptop | 2 | 40000 |
| Mobile | 3 | 15000 |
| Monitor | 4 | 10000 |

Requirement: Calculate total amount.

Formula in D2:

```excel
=B2*C2
```

Copy the formula down.

Expected result:

| Product | Quantity | Price | Total |
|---|---:|---:|---:|
| Laptop | 2 | 40000 | 80000 |
| Mobile | 3 | 15000 | 45000 |
| Monitor | 4 | 10000 | 40000 |

### 3.4 Example: Extract First Name

Assume A2 contains:

```text
Arun Kumar
```

Formula:

```excel
=TEXTBEFORE(A2," ")
```

Expected result:

```text
Arun
```

Alternative using older Excel functions:

```excel
=LEFT(A2,FIND(" ",A2)-1)
```

These formulas assume the name contains a space.

### 3.5 Example: Combine Monthly Sales

Assume three worksheets contain monthly sales:

- January
- February
- March

Using Power Query:

1. Convert each dataset into an Excel table.
2. Open Data → Get Data.
3. Import the three tables into Power Query.
4. Select Append Queries.
5. Choose Three or more tables.
6. Add January, February and March.
7. Click OK.
8. Select Close & Load.

The monthly records are combined into one dataset.

### 3.6 Example: Merge Employee and Salary Data

Employee dataset:

| Employee ID | Name |
|---|---|
| E101 | Arun |
| E102 | Priya |

Salary dataset:

| Employee ID | Salary |
|---|---:|
| E101 | 30000 |
| E102 | 35000 |

Using Power Query Merge:

1. Load both tables into Power Query.
2. Select the Employee query.
3. Choose Merge Queries.
4. Select the Salary query.
5. Match Employee ID in both tables.
6. Select the appropriate join type.
7. Expand the Salary column.
8. Load the merged result.

Expected result:

| Employee ID | Name | Salary |
|---|---|---:|
| E101 | Arun | 30000 |
| E102 | Priya | 35000 |

---

## 4. Python Integration and Real-World Business Use Cases

Python provides powerful tools for manipulating structured datasets.

Pandas is commonly used for selecting, sorting, filtering, combining and transforming Excel data.

### 4.1 Read an Excel Dataset

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

print(df.head())
```

### 4.2 Select Specific Columns

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

result = df[["Employee", "Salary"]]

print(result)
```

### 4.3 Filter Rows

Requirement: Select employees from IT.

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

it_employees = df[
    df["Department"] == "IT"
]

print(it_employees)
```

### 4.4 Filter Using Multiple Conditions

Requirement:
- Department must be IT.
- Salary must be greater than 30000.

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

result = df[
    (df["Department"] == "IT") &
    (df["Salary"] > 30000)
]

print(result)
```

Use:
- `&` for AND.
- `|` for OR.
- `~` for NOT.

Place each condition inside parentheses.

### 4.5 Sort Data

Sort employees by salary from highest to lowest.

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

sorted_df = df.sort_values(
    by="Salary",
    ascending=False
)

print(sorted_df)
```

### 4.6 Add a New Column

Calculate total sales.

```python
import pandas as pd

df = pd.read_excel("sales.xlsx")

df["Total"] = (
    df["Quantity"] * df["Price"]
)

print(df)
```

### 4.7 Rename Columns

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

df = df.rename(columns={
    "Emp_Name": "Employee",
    "Emp_Salary": "Salary"
})

print(df)
```

### 4.8 Replace Values

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

df["Department"] = df["Department"].replace({
    "IT Dept": "IT",
    "HR Dept": "HR"
})

print(df)
```

### 4.9 Combine Datasets Vertically

Use `concat()` to append rows from similar datasets.

```python
import pandas as pd

jan = pd.read_excel("january.xlsx")
feb = pd.read_excel("february.xlsx")
mar = pd.read_excel("march.xlsx")

combined = pd.concat(
    [jan, feb, mar],
    ignore_index=True
)

combined.to_excel(
    "combined_sales.xlsx",
    index=False
)
```

### 4.10 Merge Datasets Horizontally

Use `merge()` to combine datasets using a common key.

```python
import pandas as pd

employees = pd.read_excel("employees.xlsx")
salaries = pd.read_excel("salaries.xlsx")

result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="left"
)

print(result)
```

Common join types:

| Join Type | Purpose |
|---|---|
| Inner | Keeps matching records from both datasets |
| Left | Keeps all records from the left dataset |
| Right | Keeps all records from the right dataset |
| Outer | Keeps all records from both datasets |

Check whether the join key contains duplicates before merging. Duplicate keys can produce more rows than expected.

### 4.11 Group Data

Calculate department-wise average salary.

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

result = df.groupby(
    "Department"
)["Salary"].mean()

print(result)
```

### 4.12 Export Manipulated Data

```python
import pandas as pd

df = pd.read_excel("employees.xlsx")

result = df[
    df["Salary"] > 30000
]

result.to_excel(
    "high_salary_employees.xlsx",
    index=False
)
```

### 4.13 Real-World Business Use Cases

**HR Management**
- Filter employees by department.
- Sort employees by salary.
- Combine employee records.
- Create department-wise summaries.

**Sales Management**
- Combine monthly sales.
- Filter high-value transactions.
- Calculate total sales.
- Group sales by region.

**Finance**
- Combine transaction datasets.
- Filter expenses by category.
- Calculate monthly totals.
- Prepare financial reports.

**Inventory Management**
- Merge product and supplier data.
- Calculate stock value.
- Filter low-stock products.
- Group inventory by category.

**Business Analytics**
- Transform raw datasets.
- Prepare analytical datasets.
- Create summary tables.
- Generate reports for decision-making.

### 4.14 Data Manipulation in an Excel AI Agent

An Excel AI Agent can convert natural-language instructions into workbook operations.

Example user instructions:

- "Show only IT department employees."
- "Sort salary from highest to lowest."
- "Create a total amount column."
- "Combine January and February sales."
- "Merge employee and salary sheets."
- "Group sales by region."
- "Export records where sales exceed 100000."

Possible workflow:

1. Understand the user's instruction.
2. Identify the relevant worksheet and columns.
3. Read the workbook structure.
4. Determine the required manipulation operation.
5. Execute the operation using Excel tools or Pandas.
6. Validate the resulting dataset.
7. Save the output workbook.
8. Explain what changed.

The agent should distinguish between operations that modify the original workbook and operations that create a separate result.

---

## 5. Best Practices, Common Mistakes and Troubleshooting

### 5.1 Best Practices

- Keep original data backed up.
- Use clear column names.
- Check data types before manipulation.
- Use filters for targeted record selection.
- Use sorting carefully when records contain related columns.
- Verify join keys before merging datasets.
- Check row counts before and after combining data.
- Avoid unnecessary transformations.
- Validate results after every major operation.
- Save important outputs with meaningful filenames.

### 5.2 Common Mistakes

**Mistake 1: Sorting Only One Column**

Sorting one column independently can disconnect values from their corresponding records.

Solution: Select the complete dataset or use an Excel table before sorting.

**Mistake 2: Incorrect Merge Keys**

Merging datasets using unrelated or inconsistent identifiers can produce incorrect results.

Solution: Verify the common key and inspect unmatched records.

**Mistake 3: Unexpected Duplicate Rows After Merge**

Duplicate values in the join key can create multiple matching combinations.

Solution: Check key uniqueness and understand the expected relationship between datasets.

**Mistake 4: Incorrect Filter Conditions**

Using the wrong logical operator can return unexpected records.

Solution: Review the conditions and test a small sample.

**Mistake 5: Overwriting Original Data**

Directly saving transformed data over the original file can make recovery difficult.

Solution: Save a separate output file or create a backup.

**Mistake 6: Data Type Mismatch**

A numeric ID stored as text in one dataset and as a number in another can cause merge problems.

Solution: Standardize key data types before combining datasets.

### 5.3 Troubleshooting

| Problem | Possible Cause | Solution |
|---|---|---|
| Sorting produces incorrect results | Only one column was selected | Sort the complete dataset |
| Filter returns no records | Incorrect condition or inconsistent values | Check values and conditions |
| Merge produces missing values | Keys do not match | Standardize join keys |
| Merge produces too many rows | Duplicate join keys | Check key uniqueness |
| Combined data has missing columns | Different column names or structures | Standardize column names |
| Formula results are incorrect | Incorrect references or data types | Verify formulas and input values |
| Pandas cannot find a column | Incorrect column name | Inspect `df.columns` |
| Output file is missing | Incorrect output path or save operation | Verify the save path and execution |

### 5.4 Important Limitations

- Sorting can damage record relationships if the full dataset is not selected.
- Merging can create duplicate records when keys are not unique.
- Different data types can cause comparison and join problems.
- Power Query operations may need refreshing when source data changes.
- Large datasets may require more memory and processing time.
- Always verify the output before using it for business decisions.

### 5.5 Final Summary

Data Manipulation transforms and organizes raw information into useful datasets.

Key concepts:
- Selecting rows and columns.
- Sorting and filtering.
- Adding and removing columns.
- Extracting and replacing values.
- Combining and merging datasets.
- Grouping and transforming data.
- Excel operations using Power Query.
- Python operations using Pandas.
- Natural-language data manipulation through an Excel AI Agent.

Effective data manipulation is an important step in preparing reliable datasets for analysis, reporting and dashboard creation.