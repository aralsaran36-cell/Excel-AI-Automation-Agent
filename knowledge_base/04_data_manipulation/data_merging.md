# Excel Data Merging

## 1. Overview and Fundamentals

### 1.1 What Is Data Merging?

Data Merging is the process of combining information from two or more datasets using a common column or key.

It helps bring related information together into a single dataset.

Example:

Employee dataset:

| Employee ID | Name |
|---|---|
| E101 | Arun |
| E102 | Priya |
| E103 | Kumar |

Salary dataset:

| Employee ID | Salary |
|---|---:|
| E101 | 30000 |
| E102 | 35000 |
| E103 | 40000 |

After merging:

| Employee ID | Name | Salary |
|---|---|---:|
| E101 | Arun | 30000 |
| E102 | Priya | 35000 |
| E103 | Kumar | 40000 |

The Employee ID column is the common key.

### 1.2 Why Is Data Merging Important?

Data merging helps to:

- Combine information from different worksheets.
- Connect related business records.
- Create complete datasets.
- Prepare data for analysis.
- Build consolidated reports.
- Combine information from different departments.
- Support business intelligence and dashboards.

### 1.3 Data Merging vs. Data Appending

| Data Merging | Data Appending |
|---|---|
| Combines columns using a common key | Combines rows from multiple datasets |
| Usually joins related information horizontally | Usually stacks similar datasets vertically |
| Example: Employee and salary data | Example: January and February sales |
| Uses joins or lookup operations | Uses append or concatenation |

### 1.4 Important Terms

**Dataset:** A collection of structured records.

**Key Column:** A column used to identify matching records.

**Primary Key:** A column that uniquely identifies each record in a table.

**Foreign Key:** A column that refers to a key in another table.

**Join:** An operation that combines records based on matching conditions.

**Unmatched Records:** Records that do not have a corresponding match in the other dataset.

---

## 2. Core Concepts, Methods and Techniques

### 2.1 Common Data Join Types

Four commonly used join types are:

#### Inner Join

Returns only records that have matching keys in both datasets.

Example:

Dataset A:

| ID | Name |
|---|---|
| 1 | Arun |
| 2 | Priya |
| 3 | Kumar |

Dataset B:

| ID | Salary |
|---|---:|
| 1 | 30000 |
| 2 | 35000 |
| 4 | 40000 |

Inner Join result:

| ID | Name | Salary |
|---|---|---:|
| 1 | Arun | 30000 |
| 2 | Priya | 35000 |

ID 3 and ID 4 are excluded because they do not have matching records in both datasets.

#### Left Join

Returns all records from the left dataset and matching records from the right dataset.

Unmatched right-side values become missing.

Example:

| ID | Name | Salary |
|---|---|---:|
| 1 | Arun | 30000 |
| 2 | Priya | 35000 |
| 3 | Kumar | Missing |

#### Right Join

Returns all records from the right dataset and matching records from the left dataset.

Unmatched left-side values become missing.

Example:

| ID | Name | Salary |
|---|---|---:|
| 1 | Arun | 30000 |
| 2 | Priya | 35000 |
| 4 | Missing | 40000 |

#### Full Outer Join

Returns all records from both datasets.

Matching records are combined, while unmatched records are also retained.

Example:

| ID | Name | Salary |
|---|---|---:|
| 1 | Arun | 30000 |
| 2 | Priya | 35000 |
| 3 | Kumar | Missing |
| 4 | Missing | 40000 |

### 2.2 Join Type Comparison

| Join Type | Records Retained |
|---|---|
| Inner | Matching records only |
| Left | All left records and matching right records |
| Right | All right records and matching left records |
| Outer | All records from both datasets |

### 2.3 One-to-One Relationships

Each key in one dataset matches at most one key in the other dataset.

Example:

One employee ID corresponds to one employee record and one salary record.

This is a common relationship for employee master data.

### 2.4 One-to-Many Relationships

One record in the first dataset can match multiple records in the second dataset.

Example:

One customer can have multiple orders.

Customer dataset:

| Customer ID | Customer |
|---|---|
| C101 | Arun |
| C102 | Priya |

Orders dataset:

| Order ID | Customer ID | Amount |
|---|---|---:|
| O1 | C101 | 5000 |
| O2 | C101 | 3000 |
| O3 | C102 | 7000 |

After merging, Arun appears in two rows because he has two orders.

This is expected behavior.

### 2.5 Many-to-Many Relationships

Multiple records in one dataset can match multiple records in another dataset.

This can produce multiple combinations for matching keys.

Example:

If a key appears twice in Dataset A and three times in Dataset B, a many-to-many merge can produce six matching combinations for that key.

Understand the relationship before merging to avoid unexpected row multiplication.

### 2.6 Common Data Merging Methods in Excel

Excel supports several ways to merge datasets:

1. Power Query Merge.
2. XLOOKUP.
3. VLOOKUP.
4. INDEX and MATCH.
5. MATCH with INDEX.
6. Relationships in the Data Model.

Power Query Merge is especially useful for combining complete tables.

Lookup functions are useful when retrieving specific values from another table.

### 2.7 Power Query Merge vs. Append

**Merge Queries**
- Combines columns using matching keys.
- Similar to database joins.
- Useful for employee and salary datasets.

**Append Queries**
- Combines rows from multiple tables.
- Useful for monthly or regional datasets with similar columns.

### 2.8 Handling Unmatched Records

After merging, some records may not have matching values.

Possible reasons:
- Missing keys.
- Extra spaces.
- Different data types.
- Different identifier formats.
- Incorrect source data.

Useful checks:
- Identify missing values.
- Compare distinct key values.
- Standardize data types.
- Check leading and trailing spaces.
- Verify identifier uniqueness.

---

## 3. Step-by-Step Implementation with Excel Examples

### 3.1 Example: Merge Employee and Salary Tables Using Power Query

Employee table:

| Employee ID | Name |
|---|---|
| E101 | Arun |
| E102 | Priya |
| E103 | Kumar |

Salary table:

| Employee ID | Salary |
|---|---:|
| E101 | 30000 |
| E102 | 35000 |
| E103 | 40000 |

Steps:

1. Convert both datasets into Excel tables.
2. Select the Employee table.
3. Open Data → From Table/Range.
4. Load the Salary table into Power Query as well.
5. Select the Employee query.
6. Choose Home → Merge Queries.
7. Select the Salary query.
8. Select Employee ID in both tables.
9. Choose Left Outer Join.
10. Click OK.
11. Expand the Salary column.
12. Select the required fields.
13. Click Close & Load.

Expected result:

| Employee ID | Name | Salary |
|---|---|---:|
| E101 | Arun | 30000 |
| E102 | Priya | 35000 |
| E103 | Kumar | 40000 |

### 3.2 Example: Use XLOOKUP to Retrieve Salary

Assume:

- Employee IDs are in column A.
- Salary table contains IDs in column E.
- Salary values are in column F.

Formula:

```excel
=XLOOKUP(A2,E:E,F:F,"Not Found")
```

This searches for the employee ID in column E and returns the corresponding salary from column F.

XLOOKUP is available in supported modern Excel versions.

### 3.3 Example: Use VLOOKUP to Retrieve Salary

```excel
=VLOOKUP(A2,$E$2:$F$100,2,FALSE)
```

Explanation:

- A2: Lookup value.
- E2:F100: Lookup table.
- 2: Return the second column.
- FALSE: Exact match.

VLOOKUP searches in the first column of the selected table range and returns a value from a column to its right.

### 3.4 Example: Use INDEX and MATCH

```excel
=INDEX($F$2:$F$100,MATCH(A2,$E$2:$E$100,0))
```

MATCH finds the row position of the employee ID.

INDEX returns the salary from the corresponding position.

This method is useful when the return column is not necessarily to the right of the lookup column.

### 3.5 Example: Append Monthly Sales

January:

| Product | Sales |
|---|---:|
| Laptop | 50000 |
| Mobile | 30000 |

February:

| Product | Sales |
|---|---:|
| Monitor | 40000 |
| Keyboard | 10000 |

Steps:

1. Load January into Power Query.
2. Load February into Power Query.
3. Select Home → Append Queries.
4. Choose Two tables.
5. Select January and February.
6. Click OK.
7. Select Close & Load.

Expected result:

| Product | Sales |
|---|---:|
| Laptop | 50000 |
| Mobile | 30000 |
| Monitor | 40000 |
| Keyboard | 10000 |

Append combines rows rather than matching records using a key.

### 3.6 Example: Identify Unmatched Records

After merging two tables in Power Query:

1. Select the merged query.
2. Inspect the expanded columns.
3. Identify rows where the matching fields are null.
4. Check whether the keys exist in the other table.
5. Correct source values if necessary.

This helps detect missing relationships between datasets.

---

## 4. Python Integration and Real-World Business Use Cases

Pandas provides the `merge()` function to combine datasets using common columns.

### 4.1 Read Two Excel Datasets

```python
import pandas as pd

employees = pd.read_excel("employees.xlsx")
salaries = pd.read_excel("salaries.xlsx")

print(employees.head())
print(salaries.head())
```

### 4.2 Inner Join

```python
import pandas as pd

employees = pd.read_excel("employees.xlsx")
salaries = pd.read_excel("salaries.xlsx")

result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="inner"
)

print(result)
```

Only matching employee IDs are retained.

### 4.3 Left Join

```python
result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="left"
)

print(result)
```

All employee records are retained. Employees without salary matches receive missing salary values.

### 4.4 Right Join

```python
result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="right"
)

print(result)
```

All salary records are retained, including records without a matching employee.

### 4.5 Outer Join

```python
result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="outer"
)

print(result)
```

All records from both datasets are retained.

### 4.6 Merge Using Different Column Names

Sometimes the common key has different column names.

Example:
- Dataset A: Employee_ID
- Dataset B: Emp_ID

Use `left_on` and `right_on`.

```python
result = pd.merge(
    employees,
    salaries,
    left_on="Employee_ID",
    right_on="Emp_ID",
    how="left"
)

print(result)
```

### 4.7 Merge Using Multiple Columns

Sometimes records must match using more than one column.

Example:
- Employee_ID.
- Department.

```python
result = pd.merge(
    employees,
    salaries,
    on=["Employee_ID", "Department"],
    how="inner"
)

print(result)
```

All specified key columns must match.

### 4.8 Validate Merge Relationships

Pandas supports the `validate` parameter to check expected key relationships.

Example: One employee should have at most one salary record.

```python
result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="left",
    validate="one_to_one"
)
```

Common validation options:

| Option | Meaning |
|---|---|
| one_to_one | Keys must be unique in both datasets |
| one_to_many | Keys must be unique in the left dataset |
| many_to_one | Keys must be unique in the right dataset |
| many_to_many | Allows repeated keys in both datasets |

Choose the relationship that matches the business data.

### 4.9 Identify Unmatched Records

Use an indicator column to identify where records came from.

```python
result = pd.merge(
    employees,
    salaries,
    on="Employee_ID",
    how="outer",
    indicator=True
)

print(result)
```

The `_merge` column can contain:

- `left_only`
- `right_only`
- `both`

This is useful for identifying unmatched records.

### 4.10 Handle Duplicate Keys Before Merging

Check duplicate employee IDs:

```python
duplicates = employees[
    employees["Employee_ID"].duplicated(keep=False)
]

print(duplicates)
```

Inspect duplicates before removing them. Repeated keys may represent valid business records.

### 4.11 Standardize Keys Before Merging

If keys contain extra spaces, standardize them first.

```python
employees["Employee_ID"] = (
    employees["Employee_ID"].astype("string").str.strip()
)

salaries["Employee_ID"] = (
    salaries["Employee_ID"].astype("string").str.strip()
)
```

Also verify that identifiers use consistent formatting and capitalization.

### 4.12 Export the Merged Dataset

```python
result.to_excel(
    "merged_employees.xlsx",
    index=False
)
```

### 4.13 Real-World Business Use Cases

**HR Management**
- Merge employee and salary information.
- Combine attendance and employee records.
- Connect employee IDs with department details.

**Sales Management**
- Merge orders with customer information.
- Connect product IDs with product prices.
- Combine sales transactions with regional details.

**Finance**
- Merge transactions with account details.
- Connect payment records with customer information.
- Combine expense records with department data.

**Inventory Management**
- Merge stock data with product details.
- Connect supplier information with purchase records.
- Combine product IDs with pricing data.

**Business Analytics**
- Consolidate multiple business datasets.
- Build analytical datasets.
- Identify unmatched records.
- Prepare data for dashboards and reporting.

### 4.14 Data Merging in an Excel AI Agent

An Excel AI Agent can understand instructions to combine information from different worksheets or workbooks.

Example user instructions:

- "Merge employee and salary sheets using Employee ID."
- "Combine customer details with order records."
- "Keep all employee records even when salary is missing."
- "Show only matching product records."
- "Find records that do not have matching IDs."
- "Append January and February sales."

Possible workflow:

1. Understand the user's instruction.
2. Identify the source worksheets or files.
3. Read their column names and data types.
4. Identify the common key.
5. Determine the required join or append operation.
6. Check duplicate keys and expected relationships.
7. Execute the operation using Pandas or Power Query.
8. Validate row counts and unmatched records.
9. Save the resulting workbook.
10. Explain the merge operation and its results.

The agent should ask for clarification when multiple possible matching keys exist and the correct key cannot be determined safely.

---

## 5. Best Practices, Common Mistakes and Troubleshooting

### 5.1 Best Practices

- Identify the correct common key before merging.
- Standardize key data types.
- Remove unnecessary spaces from identifiers.
- Check duplicate keys.
- Understand one-to-one and one-to-many relationships.
- Choose the correct join type.
- Inspect unmatched records.
- Compare row counts before and after merging.
- Use merge validation where appropriate.
- Preserve original datasets.
- Verify the final output.

### 5.2 Common Mistakes

**Mistake 1: Using the Wrong Join Key**

Merging datasets using unrelated columns can produce incorrect results.

Solution: Identify a reliable shared identifier.

**Mistake 2: Ignoring Duplicate Keys**

Duplicate keys can multiply records during merging.

Solution: Inspect duplicates and understand the expected relationship.

**Mistake 3: Data Type Mismatch**

An ID stored as a number in one dataset and as text in another may fail to match.

Solution: Convert both key columns to a consistent type.

**Mistake 4: Unexpected Missing Values**

Some records may not have matches in the other dataset.

Solution: Inspect unmatched records and determine whether the missing relationship is expected.

**Mistake 5: Choosing the Wrong Join Type**

An inner join removes unmatched records, while a left join preserves all records from the left dataset.

Solution: Understand which records must be retained before merging.

**Mistake 6: Confusing Merge and Append**

Merge combines columns using keys. Append combines rows.

Solution: Choose the operation based on the required output structure.

### 5.3 Troubleshooting

| Problem | Possible Cause | Solution |
|---|---|---|
| Merged columns contain missing values | No matching keys | Check unmatched records |
| Output contains too many rows | Duplicate keys | Inspect key uniqueness |
| Some records disappear | Inner join used unintentionally | Use the required join type |
| Merge fails because of data types | Inconsistent key types | Standardize data types |
| Matching IDs are not recognized | Extra spaces or inconsistent formats | Clean and standardize keys |
| Column names have suffixes | Duplicate non-key column names | Rename columns before merging |
| Power Query merge returns unexpected results | Incorrect join columns or join type | Review merge configuration |
| Pandas merge raises a validation error | Key relationship differs from expectation | Inspect duplicates and correct the relationship |

### 5.4 Important Limitations

- A merge does not guarantee that records are logically related.
- Duplicate keys can create unexpected row multiplication.
- Missing or inconsistent identifiers can produce unmatched records.
- Different join types produce different results.
- Large datasets may require significant memory.
- Always validate merged records before business reporting.

### 5.5 Final Summary

Data Merging combines related datasets using common identifiers.

Key concepts:
- Common keys and relationships.
- Inner, Left, Right and Outer joins.
- One-to-one and one-to-many relationships.
- Power Query Merge and Append.
- XLOOKUP and VLOOKUP.
- Python merging using Pandas.
- Duplicate key detection.
- Unmatched record identification.
- Merge automation through an Excel AI Agent.

Correct merging is essential for building complete and reliable datasets for business analysis and reporting.