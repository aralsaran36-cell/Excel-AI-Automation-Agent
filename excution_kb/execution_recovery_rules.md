# Execution Recovery Rules

## 1. Purpose

This document defines how the Excel AI Agent detects, handles, recovers from, and reports execution failures during Excel operations.

The agent must identify the reason for an execution failure, determine whether the operation was completed partially or fully, and select an appropriate recovery strategy.

The primary objective is to recover from failures without causing duplicate operations, unintended modifications, data loss, or inconsistent workbook states.

## 2. Execution Failure Categories

The agent must recognize different types of execution failures.

| Failure Category | Description |
|---|---|
| Tool Failure | Selected execution tool fails |
| Tool Unavailable | Required tool is not available |
| Application Failure | Excel application crashes or becomes unresponsive |
| Connection Failure | Communication with the Excel application is interrupted |
| Timeout | Operation does not complete within the expected time |
| Invalid Arguments | Tool receives incorrect parameters |
| Workbook Failure | Workbook cannot be accessed or modified |
| Worksheet Failure | Target worksheet cannot be accessed |
| Range Failure | Invalid or unavailable cell range |
| Permission Failure | Operation is blocked by access restrictions |
| Formula Failure | Excel formula returns an error |
| Partial Execution | Only some requested operations succeed |
| Unexpected Output | Execution result differs from the expected result |

## 3. Failure Detection

Before attempting recovery, determine whether the execution actually failed.

Inspect the available execution result for:

- Tool execution status.
- Error messages.
- Exception details.
- Returned values.
- Workbook accessibility.
- Worksheet availability.
- Operation completion status.
- Partial execution information.

Important rules:

- Do not assume an operation failed merely because the response was delayed.
- Do not assume an operation succeeded because the tool call returned without an exception.
- Verify the actual workbook state whenever possible.
- Distinguish between execution errors and valid Excel results.
- Do not repeat an operation until its previous execution status is understood.

## 4. Error Classification and Recovery Strategy

| Error | Recovery Strategy |
|---|---|
| Invalid arguments | Correct arguments and validate again |
| Invalid worksheet | Re-identify the worksheet |
| Invalid range | Re-identify and validate the range |
| Tool unavailable | Identify another compatible available tool |
| Application disconnected | Re-establish the application connection |
| Application crashed | Restart or reinitialize the application if supported |
| Timeout | Check operation status before retrying |
| Permission denied | Identify the permission limitation |
| Workbook locked | Check workbook availability and access |
| Formula error | Inspect formula references and arguments |
| Partial execution | Identify completed and pending operations |
| Unexpected output | Validate workbook state and investigate the discrepancy |

Do not repeatedly retry an operation without identifying or addressing the failure cause.

## 5. Retry Mechanism

Retry an operation only when the failure is potentially recoverable.

### 5.1 Retry Conditions

A retry may be appropriate when:

- Temporary application disconnection occurs.
- A recoverable timeout occurs.
- A temporary tool execution error occurs.
- The application becomes available again.
- Invalid arguments have been corrected.

### 5.2 Retry Restrictions

Do not retry blindly when:

- The previous operation may have completed.
- The operation partially modified the workbook.
- The operation may create duplicate records.
- The operation may delete existing data.
- The operation may overwrite existing content.
- The execution status cannot be determined.

### 5.3 Retry Limits

- Use a limited number of retry attempts.
- Avoid infinite retry loops.
- Increase recovery effort only when the previous attempt provides useful diagnostic information.
- Stop retrying when the same failure continues.
- Report the final failure clearly.

The retry limit must be controlled by the application configuration rather than assumed to be unlimited.

## 6. Preventing Duplicate Execution

Some Excel operations are not safe to execute multiple times.

Examples:

| Operation | Potential Duplicate Effect |
|---|---|
| Append Rows | Duplicate records |
| Insert Rows | Additional unwanted rows |
| Create Worksheet | Duplicate worksheets |
| Add Formula | Repeated or misplaced formulas |
| Copy Range | Unintended repeated content |
| Send or Export Data | Duplicate output files or actions |

Before retrying these operations:

1. Inspect the current workbook state.
2. Determine whether the previous operation completed.
3. Identify which changes already exist.
4. Execute only the remaining work.
5. Verify the final result.

Prefer checking the intended postcondition rather than relying only on the previous tool response.

## 7. Partial Execution Recovery

A request may contain multiple operations, and some may succeed while others fail.

Example:

User request:

"Calculate total sales, highlight values above 50000, and create a summary sheet."

Execution status:

| Operation | Status |
|---|---|
| Calculate Total Sales | Success |
| Highlight Values | Failed |
| Create Summary Sheet | Success |

Recovery procedure:

1. Identify all requested operations.
2. Record the status of each operation.
3. Identify the failed operation.
4. Check whether successful operations modified the workbook.
5. Determine whether the failed operation can be safely retried.
6. Execute only the pending operation.
7. Verify the complete task.

Important rules:

- Never repeat already completed operations unnecessarily.
- Do not treat partial success as complete success.
- Preserve the results of successful operations.
- Report failed operations separately.
- Maintain execution status for every requested operation.

## 8. Application Crash Recovery

When the Excel application crashes or becomes unavailable:

1. Detect the application failure.
2. Stop issuing new operations to the failed application instance.
3. Identify the last known execution state.
4. Check whether the workbook is accessible.
5. Reinitialize or reconnect using supported application mechanisms.
6. Reopen the workbook if necessary.
7. Inspect whether previous changes were saved.
8. Restore the required execution context.
9. Resume only pending operations.
10. Verify the final workbook state.

Rules:

- Do not assume unsaved changes were preserved.
- Do not reopen or overwrite a workbook without checking its state.
- Avoid concurrent operations against a failed application instance.
- Do not claim recovery succeeded until the application and workbook are accessible again.

## 9. Timeout Recovery

When an operation exceeds its execution timeout:

1. Record the timeout.
2. Determine whether the operation is still running.
3. Check the workbook state.
4. Determine whether the operation completed.
5. If completed, verify its result and do not retry.
6. If still running, avoid starting a duplicate operation.
7. If terminated without completion, identify the remaining work.
8. Retry only when safe.

Do not interpret every timeout as a failed operation.

## 10. Workbook State Verification

After recovering from an execution failure, verify the workbook before continuing.

Check:

- Correct workbook is open.
- Correct worksheet is available.
- Expected data is present.
- No unintended duplicate records exist.
- No unrelated cells were modified.
- Formulas are correctly applied.
- Formatting is preserved where required.
- Worksheet structure remains valid.
- Expected changes are saved when saving is part of the operation.

Use actual workbook information rather than assuming the expected state was achieved.

## 11. Recovery from Invalid Tool Selection

If the selected execution tool cannot perform the requested operation:

1. Identify the unsupported capability.
2. Re-evaluate the operation requirements.
3. Inspect available execution tools.
4. Identify a compatible alternative.
5. Validate the alternative's capabilities.
6. Map the required parameters correctly.
7. Execute using the appropriate tool.

Important rules:

- Do not repeatedly invoke an incompatible tool.
- Do not assume every tool supports the same operations.
- Do not invent unavailable tool capabilities.
- Use an alternative only when it preserves the requested behaviour.

## 12. Formula Execution Recovery

When formula execution fails:

1. Inspect the generated formula.
2. Verify the target worksheet and cell.
3. Validate cell references.
4. Check formula argument requirements.
5. Verify data types and ranges.
6. Check Excel function compatibility.
7. Correct the identified issue.
8. Re-execute only when safe.
9. Verify the calculated result.

Distinguish between:

- Formula generation failure.
- Formula insertion failure.
- Formula calculation failure.
- Valid Excel error values.

Do not replace an unexpected formula result with an arbitrary value.

## 13. Recovery Logging

Every significant execution failure and recovery attempt should be recorded when application logging is available.

Recommended log information:

| Field | Description |
|---|---|
| Task ID | Unique execution identifier |
| Timestamp | Failure occurrence time |
| Operation | Requested Excel operation |
| Workbook | Target workbook |
| Worksheet | Target worksheet |
| Tool | Selected execution tool |
| Error Type | Identified failure category |
| Error Message | Original error details |
| Attempt Number | Recovery attempt number |
| Previous Status | Status before recovery |
| Recovery Action | Action taken |
| Final Status | Final execution status |

Rules:

- Preserve the original error message.
- Record recovery attempts separately.
- Avoid logging sensitive workbook contents unnecessarily.
- Do not replace the original error with only a generic failure message.

## 14. Recovery Decision Workflow

Follow this sequence whenever an execution failure occurs.

1. Detect the failure.
2. Capture the original error details.
3. Identify the affected operation.
4. Classify the failure.
5. Determine whether execution partially completed.
6. Inspect the current workbook state.
7. Identify the root cause where possible.
8. Select an appropriate recovery strategy.
9. Check whether retrying is safe.
10. Correct the identified issue.
11. Resume only the necessary operations.
12. Verify the final result.
13. Record the recovery outcome.
14. Report the result to the user.

## 15. Recovery Status Management

Maintain a clear status for each operation.

| Status | Meaning |
|---|---|
| Pending | Operation has not started |
| Running | Operation is currently executing |
| Success | Operation completed successfully |
| Failed | Operation failed |
| Partial | Operation completed only partially |
| Retrying | Recovery attempt is in progress |
| Skipped | Operation was intentionally not executed |
| Cancelled | Operation was cancelled |

Do not mark an operation as successful without sufficient execution evidence.

## 16. User Communication

When recovery is successful, report the completed operation accurately.

When recovery fails, explain:

- Which operation failed.
- What caused the failure, if known.
- Whether any changes were already made.
- What recovery attempts were performed.
- What remains incomplete.

When user intervention is required, request only the information or action necessary to continue.

Never claim that the workbook was successfully modified if the result has not been verified.

## 17. Important Rules

- Never retry blindly.
- Never execute duplicate operations unnecessarily.
- Always inspect execution status before retrying.
- Preserve successful operations.
- Identify partial execution.
- Verify workbook state after recovery.
- Use limited retry attempts.
- Do not hide original error information.
- Do not invent unavailable recovery capabilities.
- Avoid data loss and unintended overwriting.
- Maintain operation-level execution status.
- Stop when safe recovery is not possible.
- Report actual execution results.
- Use only available recovery and execution mechanisms.

## 18. Final Recovery Principle

The Excel AI Agent must prioritize workbook integrity over repeated execution attempts.

Every recovery decision must be based on the actual execution status, current workbook state, identified failure cause, and safety of the proposed recovery action.

The agent must resume from the last verified successful state rather than restarting the entire task unnecessarily.