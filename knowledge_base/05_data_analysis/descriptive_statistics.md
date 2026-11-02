# Descriptive Statistics in Excel

## 1. Overview and Fundamentals

### 1.1 What Is Descriptive Statistics?

Descriptive Statistics is the process of summarizing and describing the main characteristics of a dataset using numerical measurements and statistical methods.

It helps understand what the data contains without making predictions or drawing conclusions beyond the available observations.

Example:

A company has employee salary data:

| Employee | Salary |
|---|---:|
| Arun | 25000 |
| Priya | 30000 |
| Kumar | 35000 |
| Divya | 40000 |
| Ravi | 50000 |

Descriptive statistics can answer:

- What is the average salary?
- What is the middle salary?
- What is the highest salary?
- What is the lowest salary?
- How much do salaries vary?
- Which salary occurs most frequently?

### 1.2 Why Is Descriptive Statistics Important?

It helps organizations:

- Summarize large datasets.
- Understand numerical distributions.
- Identify unusual values.
- Compare different groups.
- Understand data variability.
- Prepare statistical reports.
- Support business decision-making.

### 1.3 Main Categories of Descriptive Statistics

Descriptive statistics is commonly divided into three categories.

**1. Measures of Central Tendency**

These describe the center or typical value of a dataset.

- Mean
- Median
- Mode

**2. Measures of Dispersion**

These describe how spread out the values are.

- Range
- Variance
- Standard deviation
- Interquartile range

**3. Measures of Position and Distribution**

These describe where values are located within a dataset and how the distribution behaves.

- Quartiles
- Percentiles
- Minimum
- Maximum
- Skewness
- Kurtosis

### 1.4 Sample Data vs. Population Data

**Population:** The complete group being studied.

Example: Salaries of every employee in a company.

**Sample:** A subset selected from the population.

Example: Salaries of 100 employees selected from a company with 1,000 employees.

Common Excel functions:

| Statistical Measure | Sample Function | Population Function |
|---|---|---|
| Variance | VAR.S | VAR.P |
| Standard Deviation | STDEV.S | STDEV.P |

Use sample functions when the dataset represents a sample. Use population functions when it contains the entire population of interest.

---

## 2. Core Concepts, Methods and Techniques

### 2.1 Mean

The Mean is the arithmetic average of all numerical values.

Formula:

\[
Mean = \frac{\text{Sum of all values}}{\text{Number of values}}
\]

Example:

Values: 10, 20, 30, 40, 50

\[
Mean = \frac{10+20+30+40+50}{5}=30
\]

Excel function:

```excel
=AVERAGE(A2:A6)
```

Python:

```python
mean_value = df["Sales"].mean()
```

Use cases:
- Average sales.
- Average salary.
- Average order value.
- Average production output.

**Limitation:** The mean can be strongly affected by extreme values.

### 2.2 Median

The Median is the middle value after arranging numerical values in ascending or descending order.

Example with an odd number of values:

10, 20, 30, 40, 50

Median = 30

Example with an even number of values:

10, 20, 30, 40

\[
Median = \frac{20+30}{2}=25
\]

Excel function:

```excel
=MEDIAN(A2:A6)
```

Python:

```python
median_value = df["Sales"].median()
```

Use cases:
- Typical employee salary.
- Typical property price.
- Typical customer spending.
- Understanding skewed distributions.

Median is generally less affected by extreme values than the mean.

### 2.3 Mode

The Mode is the most frequently occurring value in a dataset.

Example:

10, 20, 20, 30, 40

Mode = 20

Excel function:

```excel
=MODE.SNGL(A2:A6)
```

Python:

```python
mode_value = df["Sales"].mode()
```

Use cases:
- Most common product.
- Most frequently selected category.
- Most common customer purchase quantity.
- Most frequently occurring numerical value.

A dataset can have:
- One mode.
- Multiple modes.
- No mode when all values occur equally often.

### 2.4 Range

Range measures the difference between the maximum and minimum values.

Formula:

\[
Range = Maximum - Minimum
\]

Example:

10, 20, 30, 40, 50

\[
Range = 50-10=40
\]

Excel:

```excel
=MAX(A2:A6)-MIN(A2:A6)
```

Python:

```python
data_range = df["Sales"].max() - df["Sales"].min()
```

Range is simple but depends only on the two extreme values.

### 2.5 Variance

Variance measures how far values are spread around the mean.

A higher variance indicates greater variability.

Sample variance:

\[
s^2 = \frac{\sum (x_i-\bar{x})^2}{n-1}
\]

Population variance:

\[
\sigma^2 = \frac{\sum (x_i-\mu)^2}{N}
\]

Where:

- \(x_i\) = Individual value.
- \(\bar{x}\) = Sample mean.
- \(\mu\) = Population mean.
- \(n\) = Sample size.
- \(N\) = Population size.

Excel:

```excel
=VAR.S(A2:A100)
```

For population variance:

```excel
=VAR.P(A2:A100)
```

Python:

```python
sample_variance = df["Sales"].var()
population_variance = df["Sales"].var(ddof=0)
```

### 2.6 Standard Deviation

Standard Deviation measures the typical spread of values around the mean.

It is the square root of variance.

\[
Standard\ Deviation = \sqrt{Variance}
\]

Excel:

```excel
=STDEV.S(A2:A100)
```

Population standard deviation:

```excel
=STDEV.P(A2:A100)
```

Python:

```python
sample_std = df["Sales"].std()
population_std = df["Sales"].std(ddof=0)
```

Interpretation:

- Lower standard deviation: Values are generally closer to the mean.
- Higher standard deviation: Values are more spread out.

Standard deviation is expressed in the same units as the original data.

### 2.7 Quartiles

Quartiles divide ordered data into four sections.

- Q1: 25th percentile.
- Q2: 50th percentile, which is the median.
- Q3: 75th percentile.

Excel:

```excel
=QUARTILE.INC(A2:A100,1)
```

```excel
=QUARTILE.INC(A2:A100,2)
```

```excel
=QUARTILE.INC(A2:A100,3)
```

Python:

```python
Q1 = df["Sales"].quantile(0.25)
Q2 = df["Sales"].quantile(0.50)
Q3 = df["Sales"].quantile(0.75)
```

Quartiles are useful for understanding the distribution of values.

### 2.8 Interquartile Range (IQR)

IQR measures the spread of the middle 50% of the data.

Formula:

\[
IQR = Q3-Q1
\]

Excel:

```excel
=QUARTILE.INC(A2:A100,3)-QUARTILE.INC(A2:A100,1)
```

Python:

```python
IQR = Q3 - Q1
```

IQR is useful for identifying potential outliers.

### 2.9 Percentiles

A percentile indicates the value below which a specified percentage of observations fall.

Example:

The 90th percentile is a value at or below which approximately 90% of observations fall, according to the selected percentile method.

Excel:

```excel
=PERCENTILE.INC(A2:A100,0.90)
```

Python:

```python
percentile_90 = df["Sales"].quantile(0.90)
```

Use cases:
- Salary comparisons.
- Student performance.
- Customer spending analysis.
- Delivery time analysis.

### 2.10 Skewness

Skewness measures the asymmetry of a distribution.

**Positive Skewness:**
- Longer tail toward higher values.
- A few unusually large values may be present.

**Negative Skewness:**
- Longer tail toward lower values.
- A few unusually small values may be present.

**Approximately Zero Skewness:**
- Distribution is approximately symmetric.

Excel:

```excel
=SKEW(A2:A100)
```

Python:

```python
skewness = df["Sales"].skew()
```

Skewness helps understand the shape of a distribution.

### 2.11 Kurtosis

Kurtosis describes aspects of the tail behavior of a distribution relative to a normal distribution.

Excel:

```excel
=KURT(A2:A100)
```

Python:

```python
kurtosis = df["Sales"].kurt()
```

Excel's `KURT` and Pandas' `kurt()` report excess kurtosis.

A positive value indicates heavier tails than a normal distribution under this measure. A negative value indicates lighter tails.

Kurtosis should not be interpreted as a direct measure of how tall a chart looks.

### 2.12 Five-Number Summary

A five-number summary contains:

1. Minimum.
2. First Quartile (Q1).
3. Median (Q2).
4. Third Quartile (Q3).
5. Maximum.

It provides a compact description of the data distribution.

This summary is also useful when creating box plots.

---

## 3. Step-by-Step Implementation with Excel Examples

### 3.1 Example Dataset

Assume sales values are stored in cells A2:A11.

| Transaction | Sales |
|---|---:|
| 1 | 10000 |
| 2 | 12000 |
| 3 | 15000 |
| 4 | 18000 |
| 5 | 20000 |
| 6 | 22000 |
| 7 | 25000 |
| 8 | 28000 |
| 9 | 30000 |
| 10 | 50000 |

### 3.2 Calculate Mean

```excel
=AVERAGE(A2:A11)
```

Result: 23000

### 3.3 Calculate Median

```excel
=MEDIAN(A2:A11)
```

Result: 21000

### 3.4 Calculate Mode

```excel
=MODE.SNGL(A2:A11)
```

This dataset contains no repeated values, so Excel returns `#N/A`.

### 3.5 Calculate Minimum and Maximum

```excel
=MIN(A2:A11)
```

```excel
=MAX(A2:A11)
```

Results:
- Minimum: 10000
- Maximum: 50000

### 3.6 Calculate Range

```excel
=MAX(A2:A11)-MIN(A2:A11)
```

Result: 40000

### 3.7 Calculate Sample Variance

```excel
=VAR.S(A2:A11)
```

### 3.8 Calculate Sample Standard Deviation

```excel
=STDEV.S(A2:A11)
```

### 3.9 Calculate Quartiles

```excel
=QUARTILE.INC(A2:A11,1)
```

```excel
=QUARTILE.INC(A2:A11,2)
```

```excel
=QUARTILE.INC(A2:A11,3)
```

These return Q1, Q2 and Q3 respectively.

### 3.10 Calculate IQR

```excel
=QUARTILE.INC(A2:A11,3)-QUARTILE.INC(A2:A11,1)
```

### 3.11 Calculate the 90th Percentile

```excel
=PERCENTILE.INC(A2:A11,0.90)
```

### 3.12 Create a Descriptive Statistics Report

A report can include:

| Metric | Excel Formula |
|---|---|
| Count | `=COUNT(A2:A11)` |
| Mean | `=AVERAGE(A2:A11)` |
| Median | `=MEDIAN(A2:A11)` |
| Mode | `=MODE.SNGL(A2:A11)` |
| Minimum | `=MIN(A2:A11)` |
| Maximum | `=MAX(A2:A11)` |
| Range | `=MAX(A2:A11)-MIN(A2:A11)` |
| Variance | `=VAR.S(A2:A11)` |
| Standard Deviation | `=STDEV.S(A2:A11)` |
| Q1 | `=QUARTILE.INC(A2:A11,1)` |
| Q3 | `=QUARTILE.INC(A2:A11,3)` |

### 3.13 Use Excel Data Analysis ToolPak

Excel provides a Descriptive Statistics tool through the Analysis ToolPak.

Steps:

1. Enable the Analysis ToolPak if it is not already enabled.
2. Open the Data tab.
3. Select Data Analysis.
4. Choose Descriptive Statistics.
5. Select the input range.
6. Choose the output range.
7. Select Summary Statistics.
8. Click OK.

The output can include:
- Mean.
- Standard Error.
- Median.
- Mode.
- Standard Deviation.
- Sample Variance.
- Kurtosis.
- Skewness.
- Range.
- Minimum.
- Maximum.
- Sum.
- Count.

The exact output depends on the selected options and data.

### 3.14 Create a Box and Whisker Chart

A Box and Whisker chart helps visualize:
- Median.
- Quartiles.
- Spread.
- Potential outliers.

Steps:

1. Select the numerical dataset.
2. Open Insert.
3. Choose Insert Statistic Chart.
4. Select Box and Whisker.
5. Add a descriptive title.
6. Review the distribution and potential outliers.

---

## 4. Python Integration and Real-World Business Use Cases

Python's Pandas and NumPy libraries provide statistical functions for analyzing numerical datasets.

### 4.1 Load Data

```python
import pandas as pd

df = pd.read_excel("sales.xlsx")

print(df.head())
```

### 4.2 Generate Descriptive Statistics Automatically

```python
summary = df["Sales"].describe()

print(summary)
```

The output includes:
- Count.
- Mean.
- Standard deviation.
- Minimum.
- 25th percentile.
- 50th percentile.
- 75th percentile.
- Maximum.

### 4.3 Calculate Mean, Median and Mode

```python
mean_value = df["Sales"].mean()

median_value = df["Sales"].median()

mode_value = df["Sales"].mode()

print("Mean:", mean_value)
print("Median:", median_value)
print("Mode:", mode_value)
```

### 4.4 Calculate Range

```python
data_range = (
    df["Sales"].max() - df["Sales"].min()
)

print("Range:", data_range)
```

### 4.5 Calculate Variance and Standard Deviation

```python
sample_variance = df["Sales"].var()

sample_std = df["Sales"].std()

print("Variance:", sample_variance)
print("Standard Deviation:", sample_std)
```

### 4.6 Calculate Quartiles and IQR

```python
Q1 = df["Sales"].quantile(0.25)

Q2 = df["Sales"].quantile(0.50)

Q3 = df["Sales"].quantile(0.75)

IQR = Q3 - Q1

print("Q1:", Q1)
print("Median:", Q2)
print("Q3:", Q3)
print("IQR:", IQR)
```

### 4.7 Calculate Percentiles

```python
p90 = df["Sales"].quantile(0.90)

print("90th Percentile:", p90)
```

### 4.8 Calculate Skewness and Kurtosis

```python
skewness = df["Sales"].skew()

kurtosis = df["Sales"].kurt()

print("Skewness:", skewness)
print("Excess Kurtosis:", kurtosis)
```

### 4.9 Calculate Statistics for Multiple Columns

```python
summary = df[
    ["Sales", "Quantity"]
].describe()

print(summary)
```

This summarizes multiple numerical columns.

### 4.10 Group Descriptive Statistics

```python
summary = df.groupby("Region")["Sales"].agg(
    ["count", "mean", "median", "min", "max", "std"]
)

print(summary)
```

This compares descriptive statistics across regions.

### 4.11 Identify Potential Outliers Using IQR

```python
Q1 = df["Sales"].quantile(0.25)

Q3 = df["Sales"].quantile(0.75)

IQR = Q3 - Q1

lower_limit = Q1 - 1.5 * IQR

upper_limit = Q3 + 1.5 * IQR

outliers = df[
    (df["Sales"] < lower_limit) |
    (df["Sales"] > upper_limit)
]

print(outliers)
```

These are potential outliers, not automatically incorrect records.

### 4.12 Export Statistical Results

```python
summary = df["Sales"].describe()

summary.to_excel("descriptive_statistics.xlsx")
```

This saves the statistical summary into an Excel file.

### 4.13 Real-World Business Use Cases

**HR Analytics**
- Average employee salary.
- Median salary by department.
- Salary distribution.
- Employee age distribution.

**Sales Analytics**
- Average order value.
- Median transaction amount.
- Sales variability.
- Highest and lowest sales.

**Finance**
- Expense distribution.
- Average transaction value.
- Variability in monthly expenses.
- Potentially unusual financial transactions.

**Inventory**
- Average stock quantity.
- Minimum and maximum stock levels.
- Product demand variability.
- Distribution of inventory values.

**Customer Analytics**
- Average customer spending.
- Median purchase amount.
- Order value distribution.
- Customer purchase variability.

### 4.14 Descriptive Statistics in an Excel AI Agent

An Excel AI Agent should understand natural-language requests and select appropriate statistical methods.

Example instructions:

- "Calculate the average sales."
- "Find the median employee salary."
- "Show standard deviation by region."
- "Generate a descriptive statistics report."
- "Calculate Q1, Q2 and Q3."
- "Identify potential outliers."
- "Compare the salary distributions of departments."

Possible workflow:

1. Understand the user's request.
2. Identify the relevant worksheet and numerical columns.
3. Read and inspect the dataset.
4. Check data types and missing values.
5. Select the appropriate statistical functions.
6. Calculate the requested metrics using Pandas.
7. Validate the results.
8. Explain the statistical findings.
9. Create a report or chart if requested.
10. Save the output when required.

The agent should distinguish between sample and population statistics when the context provides enough information. If the distinction materially affects the result and is unclear, it should ask the user.

---

## 5. Best Practices, Common Mistakes and Troubleshooting

### 5.1 Best Practices

- Inspect data before calculating statistics.
- Use the correct numerical columns.
- Understand sample versus population data.
- Use median when extreme values strongly affect the mean.
- Interpret standard deviation in the context of the data.
- Review quartiles and percentiles.
- Investigate potential outliers.
- Check missing values.
- Use consistent units.
- Explain results in business language.

### 5.2 Common Mistakes

**Mistake 1: Using Mean for Highly Skewed Data**

Extreme values can significantly change the mean.

Solution: Compare mean and median before interpreting the typical value.

**Mistake 2: Confusing Sample and Population Functions**

Using the wrong variance or standard deviation function may produce a different result.

Solution: Select the function based on whether the data is a sample or the complete population.

**Mistake 3: Treating Every Outlier as an Error**

An unusual value may represent a genuine business event.

Solution: Investigate the record before changing or removing it.

**Mistake 4: Ignoring Missing Values**

Different functions may handle missing values differently.

Solution: Inspect missing values and document the treatment.

**Mistake 5: Misinterpreting Standard Deviation**

A high standard deviation does not automatically mean the data is incorrect.

Solution: Interpret variability using the business context and measurement units.

**Mistake 6: Confusing Variance and Standard Deviation**

Variance is expressed in squared units, while standard deviation uses the original units.

Solution: Use standard deviation when explaining variability in the original measurement units.

### 5.3 Troubleshooting

| Problem | Possible Cause | Solution |
|---|---|---|
| AVERAGE returns an unexpected result | Text values or incorrect range | Check the selected range and data types |
| MODE.SNGL returns #N/A | No repeated numerical value | Check whether a mode exists |
| STDEV.S returns an error | Insufficient valid numerical observations | Verify the dataset |
| Percentile result is unexpected | Incorrect percentile argument | Use a value between 0 and 1 |
| Statistical report contains unexpected values | Missing or invalid data | Inspect the source dataset |
| Python describe() omits columns | Columns are not numerical | Convert appropriate columns to numeric |
| Skewness is unexpected | Highly skewed data or outliers | Inspect the distribution |
| IQR identifies many outliers | Data distribution may be unusual | Investigate before removing values |

### 5.4 Important Limitations

- Descriptive statistics summarizes observed data; it does not automatically explain why patterns occurred.
- Mean and standard deviation can be affected by extreme values.
- Different percentile methods may produce slightly different results.
- Outlier detection methods are not proof of data errors.
- Statistical summaries should be interpreted alongside business context.
- Descriptive statistics alone cannot establish causation or guarantee future outcomes.

### 5.5 Final Summary

Descriptive Statistics helps summarize and understand numerical datasets.

Important concepts:
- Mean, median and mode.
- Range, variance and standard deviation.
- Quartiles and percentiles.
- Interquartile range.
- Skewness and kurtosis.
- Five-number summary.
- Excel statistical functions.
- Python statistical analysis using Pandas.
- Business reporting and interpretation.
- Automated statistical analysis using an Excel AI Agent.

A good statistical report should provide accurate measurements and explain what those measurements mean.