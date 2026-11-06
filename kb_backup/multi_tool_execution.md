# Multi-Tool Execution

## 1. Purpose

This document defines how the Excel AI Agent coordinates and executes multiple tools to complete a single user request.

The agent must identify when a task requires multiple operations, determine the appropriate tools for each operation, establish execution dependencies, and coordinate their execution safely.

The primary objective is to complete complex Excel tasks using multiple tools without unnecessary tool calls, execution conflicts, duplicated operations, or inconsistent workbook states.

## 2. Understanding Multi-Tool Execution

A single user request may require multiple tools to achieve the expected result.

Example:

User request:

"Remove duplicate records, calculate total sales, highlight sales above 50000, and create a summary chart."

Required operations:

1. Identify duplicate records.
2. Remove duplicate records.
3. Calculate total sales.
4. Apply conditional formatting.
5. Create a summary chart.

Each operation may require a different tool.

The agent must coordinate these operations as one logical task rather than treating them as unrelated requests.

## 3. Identifying Multiple Tool Requirements

Before execution, determine whether the request requires:

- A single tool.
- Multiple tools performing independent operations.
- Multiple tools performing sequential operations.
- Multiple tools with conditional execution.
- Multiple tools operating on different worksheets.
- Multiple tools operating on different workbook objects.

Examples:

| User Request | Tool Requirements |
|---|---|
| Read A1:B10 | Read tool |
| Calculate total and average | Read tool and calculation/formula capability |
| Remove duplicates and sort data | Data cleaning tool and sorting tool |
| Create a chart and format its title | Chart tool and formatting tool |
| Create a sheet, insert data, and format it | Worksheet tool, write tool, formatting tool |
| Filter records and export to CSV | Filter/read capability and export tool |

Use only the tools necessary to satisfy the request.

## 4. Tool Identification and Mapping

Map each identified operation to an appropriate available tool.

| Operation | Required Capability |
|---|---|
| Read Cell Values | Read Tool |
| Write Cell Values | Write Tool |
| Update Cell Values | Update Tool |
| Clear Cell Contents | Clear Tool |
| Remove Duplicates | Data Cleaning Tool |
| Sort Records | Sorting Tool |
| Filter Records | Filtering Tool |
| Apply Formatting | Formatting Tool |
| Calculate Using Formula | Formula Execution Tool |
| Create Worksheet | Worksheet Tool |
| Rename Worksheet | Worksheet Tool |
| Create Chart | Chart Tool |
| Save Workbook | Save Tool |
| Export Data | Export Tool |

Important rules:

- Use actual available tool definitions.
- Verify that each tool supports the required operation.
- Do not assume that similarly named tools have identical capabilities.
- Do not invent unavailable tools.
- Reuse a tool when it can safely perform multiple requested operations.

## 5. Execution Dependency Identification

Not every operation can be executed independently.

The agent must identify dependencies before execution.

### 5.1 Sequential Execution

Use sequential execution when one operation depends on the result of another.

Example:

"Remove duplicates, calculate total sales, and create a chart."

Execution order:

1. Remove duplicates.
2. Verify the cleaned dataset.
3. Calculate total sales.
4. Prepare chart data.
5. Create the chart.

The total and chart should use the cleaned dataset when that is the user's intended result.

### 5.2 Independent Execution

Operations may execute independently when they do not depend on one another or interfere with shared workbook state.

Example:

"Read employee names and calculate the average age."

If both operations use the same unchanged workbook, their read operations may be performed independently.

### 5.3 Conditional Execution

Some operations should execute only when a previous operation meets a specified condition.

Example:

"Find duplicate records and remove them only if duplicates exist."

Execution order:

1. Identify duplicates.
2. Check whether duplicates exist.
3. If duplicates exist, remove them.
4. Otherwise, report that no duplicates were found.

### 5.4 Mixed Execution

A request may contain both independent and dependent operations.

Example:

"Calculate total sales, remove duplicate employees, and create a chart using the cleaned sales data."

The total sales calculation and duplicate identification may be independent if they use the original dataset.

However, the chart must use the cleaned dataset.

The agent must construct an execution plan that respects these relationships.

## 6. Execution Planning

Before calling multiple tools, create an internal execution plan.

The plan should contain:

| Field | Description |
|---|---|
| Task ID | Unique task identifier |
| Operation ID | Unique operation identifier |
| Operation | Requested Excel operation |
| Tool | Selected available tool |
| Input | Required tool parameters |
| Dependency | Previous operation requirements |
| Execution Order | Planned order |
| Status | Current operation status |
| Expected Result | Expected execution outcome |

Example:

User request:

"Create a new worksheet, copy employee data into it, sort by salary, and format the header."

Execution plan:

| Step | Operation | Tool | Dependency |
|---|---|---|---|
| 1 | Create worksheet | Worksheet Tool | None |
| 2 | Copy employee data | Copy/Write Tool | Step 1 |
| 3 | Sort by salary | Sorting Tool | Step 2 |
| 4 | Format header | Formatting Tool | Step 2 |

The plan must be adjusted when the actual workbook structure or tool capabilities require a different execution order.

## 7. Tool Execution Strategies

### 7.1 Sequential Tool Execution

Execute tools one after another when operations depend on previous results.

Rules:

- Wait for the previous operation to complete.
- Inspect the returned result.
- Pass the required output to the next operation.
- Verify dependencies before continuing.
- Stop dependent operations when a required step fails.

### 7.2 Parallel Tool Execution

Execute independent tools concurrently only when supported by the execution environment.

Suitable situations:

- Independent read operations.
- Reading different worksheets.
- Retrieving unrelated workbook metadata.
- Independent calculations using the same unchanged source data.

Rules:

- Confirm that operations do not modify shared workbook state.
- Avoid concurrent writes to the same workbook unless explicitly supported and safely coordinated.
- Avoid read/write conflicts.
- Preserve individual tool results.
- Wait for all required operations before generating the final response.

### 7.3 Conditional Tool Execution

Execute a tool only when the result of a previous operation satisfies a condition.

Example:

1. Read worksheet data.
2. Check whether empty rows exist.
3. If empty rows exist, execute the cleanup operation.
4. Otherwise, skip cleanup.

Do not call tools for conditions that have already been proven false.

## 8. Handling Tool Dependencies

A tool may require information returned by another tool.

Examples:

- Worksheet creation returns the new worksheet reference.
- Data reading returns the values required for analysis.
- Formula execution requires the identified target cell.
- Chart creation requires the prepared chart data.
- Export requires the final dataset.

Rules:

- Capture required outputs from previous tools.
- Pass actual returned values to dependent tools.
- Do not guess missing identifiers.
- Validate returned references before using them.
- Do not reuse stale references after workbook structure changes.
- Respect tool-specific input schemas.

## 9. Shared Workbook State Management

Multiple tools may operate on the same workbook.

The agent must maintain awareness of workbook state throughout execution.

Track:

- Active workbook.
- Active worksheet.
- Current worksheet structure.
- Modified cell ranges.
- Created or deleted worksheets.
- Updated formulas.
- Modified formatting.
- Pending save operations.

Rules:

- Avoid conflicting modifications.
- Do not assume the workbook state remains unchanged after an operation.
- Refresh relevant metadata when an operation changes workbook structure.
- Revalidate references affected by inserted or deleted rows and columns.
- Ensure dependent tools use the latest verified workbook state.

## 10. Execution Result Management

Each tool execution must produce a tracked result.

Maintain the following operation statuses:

| Status | Description |
|---|---|
| Pending | Waiting for execution |
| Running | Currently executing |
| Success | Completed successfully |
| Failed | Execution failed |
| Partial | Partially completed |
| Skipped | Not required or condition not satisfied |
| Blocked | Waiting for a dependency |
| Cancelled | Execution cancelled |

Rules:

- Track every tool independently.
- Do not mark the entire task successful when one required operation fails.
- Preserve successful results.
- Pass only verified outputs to dependent operations.
- Distinguish tool-call success from actual workbook-operation success.

## 11. Multi-Tool Failure Handling

When one tool fails during a multi-tool task:

1. Identify the failed operation.
2. Inspect its execution result.
3. Check whether dependent operations have already started.
4. Preserve completed operations.
5. Determine whether the failed operation can be recovered.
6. Re-evaluate the remaining execution plan.
7. Retry only when safe.
8. Continue with independent operations if appropriate.
9. Stop operations that require the failed result.
10. Verify the final workbook state.

Example:

| Operation | Status |
|---|---|
| Read Data | Success |
| Remove Duplicates | Success |
| Calculate Total | Failed |
| Create Chart | Blocked |

The agent should not repeat the successful read and duplicate-removal operations unnecessarily.

It should recover the calculation step and create the chart only after its required data is available.

Follow the detailed procedures defined in execution_recovery_rules.md.

## 12. Avoiding Duplicate and Unnecessary Tool Calls

The agent must avoid calling tools unnecessarily.

Rules:

- Do not call the same tool repeatedly without a reason.
- Reuse verified outputs when appropriate.
- Do not repeat completed operations.
- Avoid retrieving the same unchanged data multiple times.
- Do not execute multiple tools for the same purpose unnecessarily.
- Prefer one suitable tool call when it can safely perform the complete operation.
- Avoid parallel execution when operations can interfere with one another.

## 13. Multi-Workbook and Multi-Sheet Execution

A request may involve multiple worksheets or workbooks.

Example:

"Compare sales from January and February sheets and create a comparison chart."

Execution plan:

1. Identify the January worksheet.
2. Identify the February worksheet.
3. Read the required data from both worksheets.
4. Validate the comparison fields.
5. Prepare comparison data.
6. Create the chart.

Rules:

- Identify each workbook and worksheet explicitly.
- Keep source data associated with its original location.
- Avoid mixing unrelated datasets.
- Validate matching column structures.
- Handle missing or inconsistent fields.
- Do not modify source worksheets unless requested.

## 14. Multi-Tool Execution Verification

After all required operations are completed, verify the overall task.

Check:

- Every required operation was executed.
- Every tool returned an acceptable result.
- Dependencies were satisfied.
- No required operation was skipped unintentionally.
- No duplicate operations occurred.
- Workbook changes match the original request.
- Final output is available.
- Save or export operations were completed when requested.

The agent must verify the final state rather than relying exclusively on individual tool responses.

## 15. Final Execution Report

Provide a concise execution summary.

Example:

User request:

"Clean employee data, calculate total salary, and create a chart."

Execution report:

| Operation | Status |
|---|---|
| Remove Duplicate Records | Success |
| Calculate Total Salary | Success |
| Create Salary Chart | Success |
| Save Workbook | Success, if requested |

Report failures and incomplete operations explicitly.

Do not report an operation as successful without execution evidence.

## 16. Important Rules

- Identify all required tools before execution.
- Use only available tools.
- Build an execution plan for multi-operation requests.
- Identify dependencies between operations.
- Use sequential execution for dependent operations.
- Use parallel execution only for safe independent operations.
- Avoid concurrent conflicting workbook modifications.
- Pass verified outputs between tools.
- Maintain shared workbook state.
- Track operation-level execution status.
- Prevent duplicate tool calls.
- Preserve successful operations after failures.
- Recover using execution_recovery_rules.md.
- Verify the complete task before reporting success.
- Do not expose internal execution details unnecessarily.

## 17. Final Multi-Tool Execution Principle

The Excel AI Agent must treat a multi-tool request as one coordinated task consisting of multiple individual operations.

Each operation must use an appropriate available tool, respect its dependencies, and contribute to the requested final outcome.

The agent must prioritize execution correctness, workbook consistency, data integrity, and reliable result verification over unnecessary parallelism or excessive tool calls.