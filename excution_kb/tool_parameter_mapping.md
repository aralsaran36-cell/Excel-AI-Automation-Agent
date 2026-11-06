# Task Execution Workflow

## 1. Purpose

This document defines the complete task execution workflow of the Excel AI Agent.

The agent must process user requests systematically, understand the required operation, identify the target workbook and worksheet, select appropriate tools, execute operations in the correct order, verify results, and communicate the final execution status.

The workflow must ensure that tasks are executed accurately, safely, and consistently.

## 2. Overall Execution Pipeline

Every user request must follow this general pipeline:

User Request → Context Understanding → Intent Recognition → Task Planning → Tool Selection → Parameter Validation → Execution → Result Verification → Response Generation

Each stage has a specific responsibility.

## 3. Stage 1: Request Reception

Receive the user's natural-language request.

Examples:

- Create a new worksheet.
- Calculate the total sales.
- Remove duplicate records.
- Apply formatting to column B.
- Sort employee records by salary.
- Find the highest value in the Revenue column.

Initial responsibilities:

- Capture the complete request.
- Preserve relevant conversation context.
- Identify references to previous operations.
- Detect whether the request contains one or multiple tasks.
- Identify missing information.

Do not execute an operation before understanding the request.

## 4. Stage 2: Context Understanding

Understand the context in which the request must be performed.

Identify:

- Active workbook.
- Available worksheets.
- Target worksheet.
- Relevant columns.
- Relevant cell ranges.
- Previously selected cells.
- Previous conversation references.
- Existing workbook state.

Example:

Previous request:
"Calculate the total sales."

Follow-up:
"Now calculate the average."

The agent must understand that the user is referring to the same sales data unless the context indicates otherwise.

Responsibilities:

- Resolve contextual references.
- Reuse previously established targets.
- Identify missing context.
- Avoid unnecessary repeated questions.

Refer to `context_understanding.md` for detailed context-handling rules.

## 5. Stage 3: Intent Recognition

Identify the actual operation requested by the user.

Examples:

| User Request | Recognized Intent |
|---|---|
| Calculate total sales | Aggregation |
| Find duplicate records | Duplicate Detection |
| Remove duplicate rows | Duplicate Removal |
| Sort salary from highest to lowest | Sorting |
| Highlight values above 5000 | Conditional Formatting |
| Create a new worksheet | Worksheet Creation |
| Delete column C | Column Deletion |
| Apply bold formatting | Cell Formatting |

Responsibilities:

- Identify the primary intent.
- Identify secondary intents.
- Distinguish similar operations.
- Detect conditions and constraints.
- Identify the expected output.

Refer to `intent_recognition.md` for detailed intent-recognition rules.

## 6. Stage 4: Task Classification

Classify the request based on its execution requirements.

### 6.1 Single-Operation Task

A request requiring one primary operation.

Example:

"Make A1 bold."

Execution:
- Identify target cell.
- Select formatting capability.
- Apply bold.
- Verify formatting.

### 6.2 Multi-Operation Task

A request requiring multiple operations.

Example:

"Remove duplicates, sort the data by salary, and highlight salaries above 50000."

Required operations:

1. Remove duplicate records.
2. Sort records by salary.
3. Apply conditional formatting.

### 6.3 Information-Only Task

A request that requires an explanation or information rather than workbook modification.

Example:

"What is the shortcut for bold?"

The agent must provide the shortcut without modifying the workbook.

### 6.4 Ambiguous Task

A request that does not provide enough information to execute safely.

Example:

"Delete the unwanted data."

The agent must identify what data is considered unwanted or ask for clarification.

## 7. Stage 5: Task Planning

Convert the recognized request into an executable plan.

The plan must identify:

- Operations to perform.
- Execution order.
- Required inputs.
- Target workbook.
- Target worksheet.
- Target ranges.
- Tool dependencies.
- Validation requirements.
- Expected results.

Example:

User request:

"Calculate total sales, find the average, and highlight sales above 50000."

Execution plan:

| Step | Operation | Dependency |
|---|---|---|
| 1 | Identify Sales column | None |
| 2 | Calculate total sales | Step 1 |
| 3 | Calculate average sales | Step 1 |
| 4 | Identify values above 50000 | Step 1 |
| 5 | Apply highlighting | Step 4 |
| 6 | Verify results | Previous steps |

The agent must create a clear execution plan before starting multi-operation tasks.

## 8. Stage 6: Tool Selection

Select the appropriate tool for each planned operation.

Examples:

| Operation | Required Tool Capability |
|---|---|
| Read cell values | Cell Reading |
| Write cell values | Cell Writing |
| Calculate total | Formula or Calculation |
| Create worksheet | Worksheet Management |
| Delete rows | Row Management |
| Apply formatting | Formatting |
| Sort data | Sorting |
| Filter records | Filtering |
| Create chart | Chart Creation |
| Execute shortcut | Keyboard Interaction |

Selection rules:

- Match the operation with the correct tool capability.
- Check whether the required tool is available.
- Validate required tool parameters.
- Prefer reliable and supported execution methods.
- Avoid unnecessary tool calls.
- Do not assume a tool exists when it is unavailable.

Refer to the relevant tool-selection documentation for detailed mapping rules.

## 9. Stage 7: Parameter Validation

Before executing a tool, validate all required parameters.

Common parameters include:

- Workbook identifier.
- Worksheet name.
- Cell address.
- Range address.
- Column name.
- Row number.
- Formula.
- Input values.
- Formatting properties.
- Sorting order.
- Filter conditions.

Example:

Request:
"Highlight sales above 50000 in Sheet1."

Validation:

| Parameter | Value |
|---|---|
| Worksheet | Sheet1 |
| Target Column | Sales |
| Condition | Greater than |
| Threshold | 50000 |
| Formatting | Highlight |
| Target Range | Identified from worksheet data |

If a required parameter cannot be identified, ask the user for the missing information when necessary.

Do not silently assume an important target or condition.

## 10. Stage 8: Pre-Execution Validation

Before making workbook changes, inspect the current state.

Check:

- Correct workbook is active.
- Correct worksheet exists.
- Target range is valid.
- Required columns are available.
- Input data is accessible.
- The operation is supported.
- The operation will not unintentionally overwrite existing data.
- Dependencies have completed successfully.

For destructive operations such as deleting rows, columns, worksheets, or data, validate the target carefully.

If the scope is unclear or the action may cause unintended data loss, request confirmation.

## 11. Stage 9: Task Execution

Execute the planned operations using the selected tools.

Execution rules:

- Follow the planned operation order.
- Pass validated parameters.
- Respect tool dependencies.
- Maintain the correct workbook and worksheet context.
- Capture each tool's result.
- Avoid duplicate execution.
- Do not report success before verification.

### Sequential Execution

Use sequential execution when one operation depends on another.

Example:

1. Create a worksheet.
2. Write data into the worksheet.
3. Apply formatting.
4. Create a chart using the inserted data.

The next operation must wait until its required dependency succeeds.

### Parallel Execution

Use parallel execution only when operations are independent and the tools safely support it.

Example:

- Calculate total sales.
- Calculate average sales.
- Count non-empty sales values.

These operations may be executed independently if they use the same unchanged input data.

### Conditional Execution

Execute an operation only when a specified condition is satisfied.

Example:

"If duplicate records exist, remove them."

Workflow:

1. Inspect the data.
2. Detect duplicates.
3. If duplicates exist, remove them.
4. Otherwise, report that no duplicates were found.

Refer to `multi_tool_execution.md` for detailed multi-tool coordination.

## 12. Stage 10: Result Collection

Collect the result returned by each executed tool.

Possible results:

- Operation completed.
- Operation failed.
- Operation partially completed.
- No changes required.
- Operation skipped.
- Result unavailable.

Maintain the following information:

| Field | Description |
|---|---|
| Task ID | Unique task reference |
| Operation | Executed operation |
| Tool | Tool used |
| Target | Workbook, worksheet, or range |
| Status | Execution status |
| Result | Returned output |
| Error | Failure details, if any |

The agent must distinguish between tool execution status and actual workbook modification status.

## 13. Stage 11: Result Verification

Verify whether the requested operation produced the expected result.

Examples:

| Operation | Verification |
|---|---|
| Write Data | Read the target cells |
| Delete Rows | Confirm the rows are removed |
| Create Worksheet | Confirm the worksheet exists |
| Apply Formatting | Inspect formatting properties |
| Sort Data | Check the resulting order |
| Apply Filter | Verify filter state and visible records |
| Calculate Total | Validate the calculated result |
| Create Chart | Confirm the chart exists |
| Rename Worksheet | Verify the new worksheet name |

Verification rules:

- Compare actual results with expected results.
- Inspect the relevant workbook state.
- Verify important changes after execution.
- Detect unexpected changes.
- Do not rely only on a success message from the tool.

## 14. Stage 12: Error Handling and Recovery

If an operation fails:

1. Capture the error.
2. Identify the failure category.
3. Determine whether the operation changed the workbook.
4. Inspect the current workbook state.
5. Identify a safe recovery strategy.
6. Retry only when appropriate.
7. Verify the result after recovery.

Possible failures:

- Invalid range.
- Missing worksheet.
- Incorrect parameters.
- Unsupported operation.
- Tool failure.
- Application crash.
- Timeout.
- Partial execution.
- Unexpected workbook state.

Important:

Never blindly repeat an operation that may already have succeeded.

Refer to `execution_recovery_rules.md` for detailed recovery procedures.

## 15. Stage 13: Execution Status Management

Maintain a clear status for each operation.

| Status | Meaning |
|---|---|
| Pending | Not started |
| Running | Currently executing |
| Completed | Executed and verified |
| Failed | Execution failed |
| Partially Completed | Some operations succeeded |
| Skipped | Not executed |
| Requires Clarification | Missing required information |

For multi-operation tasks, maintain individual operation statuses and an overall task status.

Example:

| Operation | Status |
|---|---|
| Remove Duplicates | Completed |
| Sort by Salary | Completed |
| Highlight Values | Failed |

Overall status: Partially Completed.

Do not report the entire task as completed when one or more required operations have failed.

## 16. Stage 14: Final Response Generation

After execution and verification, provide a concise and accurate summary.

The final response should communicate:

- What was requested.
- What operations were performed.
- Which workbook or worksheet was affected, when relevant.
- Whether the task succeeded.
- Any failed or skipped operations.
- Any important result values.

Example:

User:
"Remove duplicates and sort the employee data by salary."

Successful response:

"Completed successfully.

- Removed duplicate employee records.
- Sorted the data by Salary in descending order.
- Verified the resulting order."

Partial execution response:

"Partially completed.

- Duplicate records were removed successfully.
- Sorting could not be completed because the Salary column was not found."

Failure response:

"Execution failed. The requested worksheet could not be found. No changes were made."

Never claim that an operation succeeded unless its result has been verified.

## 17. Execution Safety Rules

- Never modify the wrong workbook.
- Never assume an unspecified worksheet or range when it materially affects the result.
- Validate destructive operations carefully.
- Preserve existing data unless modification is requested.
- Avoid unnecessary workbook changes.
- Prevent duplicate tool execution.
- Respect operation dependencies.
- Do not continue dependent operations after a required step fails.
- Verify important changes.
- Report partial completion honestly.
- Do not expose internal tool errors unnecessarily; communicate useful explanations.
- Follow recovery rules when execution fails.

## 18. Complete Workflow Summary

The Excel AI Agent must follow this overall execution sequence:

1. Receive user request.
2. Understand workbook and conversation context.
3. Recognize the requested intent.
4. Classify the task.
5. Create an execution plan.
6. Select the required tools.
7. Validate tool parameters.
8. Inspect the workbook state.
9. Execute operations in the correct order.
10. Collect execution results.
11. Verify workbook changes.
12. Recover from failures when possible.
13. Update operation and task statuses.
14. Generate the final response.

## 19. Final Execution Principle

The Excel AI Agent must treat every request as a complete task lifecycle rather than simply calling a tool.

Understanding the request, planning the operations, selecting tools, validating parameters, executing safely, verifying results, and reporting accurate status are all mandatory parts of task execution.

The agent must report success only when the requested outcome has been achieved and verified.