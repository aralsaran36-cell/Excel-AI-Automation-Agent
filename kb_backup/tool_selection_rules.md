# Tool Selection Rules

## 1. Purpose

This document defines how the Excel AI Agent should select the correct MCP tool based on the user's natural-language request.

The agent must understand the user's intention, inspect the available tools, identify the tool that can perform the requested operation, and select it for execution.

Tool selection must be based on actual tool descriptions, capabilities, input parameters, and execution requirements.

The agent must never select a tool based only on its name.

## 2. Tool Selection Principles

Follow these rules when selecting an Excel tool:

1. Understand what the user wants to achieve.
2. Identify the required Excel operation.
3. Retrieve the available MCP tool descriptions and schemas.
4. Find tools whose capabilities match the requested operation.
5. Check the required input parameters.
6. Verify that the selected tool supports the requested operation.
7. Select the most appropriate tool.
8. If multiple tools are required, determine their execution order.
9. If no available tool supports the request, do not invent a tool.
10. Ask the user for clarification when essential information is missing.

## 3. Natural Language to Operation Mapping

The agent must identify the intended operation from different natural-language expressions.

### 3.1 Calculation Requests

Examples:

- Calculate total sales.
- Find the sum of units.
- What is the average salary?
- Find the highest sales value.
- Calculate the minimum price.
- Count the number of employees.

Required operation:

Identify the appropriate calculation, such as SUM, AVERAGE, MAX, MIN, or COUNT.

Tool selection:

Select an available tool that supports the requested calculation. If formula execution is required, select the appropriate formula tool based on its actual description and schema.

### 3.2 Data Analysis Requests

Examples:

- Analyze the sales column.
- Find unusual values.
- Show the distribution of products.
- Compare two datasets.
- Identify missing values.
- Find duplicate records.

Required operation:

Identify the analysis method required by the request.

Tool selection:

Select the available analysis tool whose documented functionality matches the requested operation.

### 3.3 Sorting and Filtering Requests

Examples:

- Sort sales from highest to lowest.
- Show only North region records.
- Filter employees with salaries above 50000.
- Display products belonging to a specific category.

Required operation:

Identify the sorting or filtering condition.

Tool selection:

Select the actual sorting or filtering tool that supports the requested condition and data range.

### 3.4 Worksheet Requests

Examples:

- Create a new worksheet.
- Rename the current sheet.
- Delete a worksheet.
- Duplicate an existing worksheet.
- List all worksheets.

Required operation:

Identify the requested worksheet operation.

Tool selection:

Select the corresponding available worksheet tool.

### 3.5 Formatting Requests

Examples:

- Make the header bold.
- Change the background color.
- Highlight duplicate values.
- Format sales as currency.
- Adjust column width.

Required operation:

Identify the formatting operation and its target range.

Tool selection:

Select the available formatting tool that supports the requested change.

### 3.6 Chart and Dashboard Requests

Examples:

- Create a sales chart.
- Generate a regional sales dashboard.
- Show monthly revenue trends.
- Create a KPI summary.
- Refresh an existing dashboard.

Required operation:

Identify whether the user wants a chart, dashboard, KPI calculation, or dashboard modification.

Tool selection:

Select the available chart or dashboard tool based on its documented capabilities.

## 4. Tool Description and Schema Verification

Before selecting a tool, the agent must inspect its actual MCP definition.

Check the following:

- Tool name.
- Tool description.
- Supported functionality.
- Required parameters.
- Optional parameters.
- Parameter data types.
- Expected input format.
- Return information.
- Known limitations.

The tool description explains when and why a tool should be used.

The schema defines how the tool must be called.

Both must be considered before execution.

Never assume that a tool supports an operation simply because its name appears relevant.

## 5. Tool Selection Decision Process

Use the following decision process:

Step 1: Understand the user request.

Step 2: Identify the primary Excel operation.

Step 3: Determine the required input data.

Step 4: Inspect the workbook context when necessary.

Step 5: Retrieve available MCP tool definitions.

Step 6: Compare tool capabilities with the requested operation.

Step 7: Verify the selected tool's parameters.

Step 8: Select the most suitable tool.

Step 9: Pass the selected tool and its requirements to the execution planner.

Step 10: Execute and validate the result.

## 6. Multiple Tool Selection

Some user requests require multiple tools.

Example:

User request:

"Create a sales summary by region and generate a chart."

Possible workflow:

1. Inspect the workbook and worksheet.
2. Identify the Region and Sales columns.
3. Calculate sales by region using an available aggregation tool.
4. Create a chart using the available chart tool.
5. Verify the generated chart.

The agent must select tools based on actual dependencies.

Do not assume that every task requires multiple tools. Use only the tools necessary to complete the request.

## 7. Tool Selection Priority

When multiple tools appear relevant, apply these priorities:

1. Prefer a tool that directly performs the requested operation.
2. Prefer a tool whose documented functionality closely matches the request.
3. Verify that the tool supports all required conditions.
4. Avoid unnecessary intermediate operations.
5. Use multiple tools only when one tool cannot complete the entire task.
6. Do not select tools with incompatible parameters or unsupported capabilities.

## 8. Handling Missing or Unsupported Tools

If no available tool supports the requested operation:

1. Do not invent a tool name.
2. Do not call an unrelated tool.
3. Check whether the operation can be completed using another documented tool.
4. If an alternative is available, use it only when its capabilities genuinely support the task.
5. If no suitable option exists, inform the user that the operation is currently unsupported.

If the tool exists but required information is missing, request the necessary information before execution.

## 9. Avoid Incorrect Tool Selection

The agent must avoid the following mistakes:

- Selecting tools based only on similar names.
- Confusing analysis tools with modification tools.
- Selecting a chart tool for a calculation-only request.
- Selecting a formatting tool for a data-cleaning request.
- Assuming unsupported parameters exist.
- Calling multiple tools unnecessarily.
- Repeating the same tool call without checking its previous result.
- Selecting a tool without verifying its documented functionality.

## 10. Tool Selection Examples

### Example 1: Calculate Total Sales

User request:

"Calculate total sales."

Intent:

Calculate the sum of sales values.

Required capability:

A tool that supports the requested calculation or formula execution.

Selection:

Choose the appropriate available MCP tool after checking its description and schema.

### Example 2: Find Outliers

User request:

"Find unusual values in the Sales column."

Intent:

Identify potential outliers.

Required capability:

Outlier detection.

Selection:

Choose the available outlier analysis tool if its documented functionality supports this operation.

### Example 3: Create a Worksheet

User request:

"Create a new worksheet called Summary."

Intent:

Create a worksheet.

Required capability:

Worksheet creation.

Selection:

Choose the available worksheet creation tool and verify its actual parameters.

### Example 4: Create a Dashboard

User request:

"Create a dashboard for monthly sales."

Intent:

Create a dashboard using monthly sales information.

Required capabilities:

- Inspect sales and date information.
- Prepare monthly sales data if required.
- Generate the dashboard.

Selection:

Choose the available dashboard and supporting data tools based on their actual capabilities.

## 11. Final Validation Rules

Before passing a selected tool to the executor, verify:

- Does the tool support the requested operation?
- Are the required parameters available?
- Is the tool compatible with the current workbook context?
- Are all required inputs identified?
- Is the execution sequence correct?
- Does the operation require user clarification?

Only after these checks should the selected tool be passed to the execution stage.

## 12. Important Rules

- Actual MCP tool definitions are the primary source for tool selection.
- Tool descriptions determine the intended use of each tool.
- Tool schemas determine the valid arguments.
- Workbook inspection provides the actual sheet, column, and range information.
- RAG knowledge provides additional execution guidance.
- Never invent tools, parameters, or capabilities.
- Never assume that a tool succeeded before checking its result.
- Always validate the execution outcome.

The objective is to select the correct available Excel tool for every user request and ensure that the selected tool matches the actual task requirements.