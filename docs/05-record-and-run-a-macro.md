# Task 04 — Record and Run a Macro

**Status:** Initial Record and Run a Macro exercise complete — learner-confirmed. The learner reports successfully completing the recording, test run, and VBA inspection steps. The exact macro name and action were not supplied, so this guide does not claim a specific macro was recorded. The learner has also confirmed adding a Form Control button for that first recorded macro. The exact macro name and recorded action remain undocumented. The two sorting macros, their individual buttons, successful tests, and readable PDF evidence are complete; the Forage upload/submission has also been completed; the activity now displays its completion page and certificate dated November 4, 2026. See [Task 04 — Sorting Macros and Buttons](06-sort-macros-and-buttons.md).

## 1. Workbook and safe working copy

- Supplied workbook: [Account_Sales_Data_for_Automation.xlsx](../material/04-vba-and-marcos/Account_Sales_Data_for_Automation.xlsx)
- Repository folder name: `material/04-vba-and-marcos/` (the existing folder spelling is retained to avoid breaking links).
- Use the clean automation workbook for Task 04. Do not use the completed Task 03 dashboard workbook for this exercise.
- Open the downloaded workbook in desktop Microsoft Excel.
- Use **File → Save As → Excel Macro-Enabled Workbook (*.xlsm)** and save a separate working copy, for example `Account_Sales_Data_for_Automation_Practice.xlsm`.
- Keep the supplied `.xlsx` as the untouched source.

The learner should keep the working `.xlsm` in a local practice location. Do not publish a macro-enabled copy or redistribute the supplied workbook unless the Forage materials' sharing terms permit it. The learner has confirmed that the initial recording, test, and code-inspection exercise was completed successfully.

## 2. Enable the Developer tab

In Excel for Windows:

1. Select **File → Options**.
2. Open **Customize Ribbon**.
3. Under **Main Tabs**, check **Developer**.
4. Select **OK**.

Use desktop Excel for this VBA exercise. The Developer tab in Excel for the web does not provide the desktop VBA recording and editing workflow.

## 3. Inspect the workbook before recording

1. Identify the worksheet containing the account-level data.
2. Confirm the header row and the first data row. Do not assume the worksheet name or hard-code a range until you have checked the workbook.
3. Locate the last used header cell. If the headers begin in A1, select A1 and use `Ctrl + Shift + Right` to select the contiguous header cells; verify the selection ends at the last header and does not include blank columns.
4. Keep the source data intact. The first macro is a low-risk practice macro that formats the existing header row.

## 4. Record a first practice macro

This exercise asks you to record actions, so use Excel's recorder rather than pasting pre-written VBA.

Suggested macro name: `FormatSalesHeaders`

Suggested description: `Apply consistent formatting to the account data headers`

1. Select the verified header cells only.
2. Select **Developer → Record Macro**.
3. Enter `FormatSalesHeaders` as the macro name. Do not use spaces.
4. Set **Store macro in** to **This Workbook**.
5. Add the description above, then select **OK**.
6. Apply a clear, repeatable header style through the ribbon:
   - Bold font.
   - Dark navy fill.
   - White font color.
   - Center alignment.
   - Wrap Text.
   - AutoFit the selected columns, checking that none becomes excessively wide.
7. Select **Developer → Stop Recording**.

Do not edit, sort, delete, or overwrite account records while recording this first formatting macro.

If the Forage page or your workbook specifies a different action, follow that requirement instead and record the actual specified action. The header-formatting workflow is a safe practice suggestion for the generic instructions supplied here, not a claim that Forage mandates this exact macro.

## 5. Test the recorded macro

1. Select the header range again.
2. Use **Home → Clear → Clear Formats** to remove only the formatting you just applied. Do not choose Clear All.
3. Select **Developer → Macros**.
4. Choose `FormatSalesHeaders` and select **Run**.
5. Confirm that the header formatting is reapplied and the header text/data remain unchanged.
6. Save the workbook with **Ctrl + S**.

A useful test is whether the macro repeats the same result without manual formatting. If it selects the wrong range or affects unrelated cells, stop and correct the recorded action before proceeding.

## 6. Inspect the generated VBA

1. Select **Developer → Macros**.
2. Choose `FormatSalesHeaders`.
3. Select **Edit**.
4. In the Visual Basic Editor, locate the new procedure, usually in a standard module such as Module1.
5. Read the generated statements and compare them with the steps you performed.

The recorder may generate code using `Select`, `Activate`, and `Selection`. That is normal for a beginner recording exercise. Later, learn how to simplify it to explicit worksheet and range references.

Do not replace the Task 04 recorded macro with the Task 03 dashboard styling module. They are separate learning exercises in separate workbooks.

## 7. Save and evidence

- Save the working file as `.xlsm`.
- Confirm `FormatSalesHeaders` appears in **Developer → Macros** after closing and reopening the workbook.
- Confirm the macro can run from the Macro dialog.
- Capture evidence of the macro dialog, successful formatting result, and the generated procedure in the VBA editor, while hiding any personal information.
- Record the actual macro name, selected range, workbook sheet name, test result, and any errors in the learning log.

## 8. Completion record

The learner confirms successful completion of the initial recording exercise, including running the macro and inspecting the generated VBA. The exact recorded macro name/action has not been captured here. The XLSM copy and original workbook should remain separate.

The learner has completed the sorting-and-buttons workflow described in [Task 04 — Sorting Macros and Buttons](06-sort-macros-and-buttons.md). The remaining action is to upload the macro-enabled workbook and evidence PDF to Forage and submit them.

## Completion checklist

- [x] Initial macro-recording exercise completed — learner-confirmed.
- [x] Recorded macro tested successfully — learner-confirmed.
- [x] Generated VBA inspected in the Visual Basic Editor — learner-confirmed.
- [x] Add a Form Control button for the initial recorded macro — learner-confirmed.
- [x] Complete the two sorting macros and their buttons — learner-confirmed.
- [x] Complete final sorting-button tests and requested PDF evidence — learner-confirmed.
- [x] Upload the XLSM and PDF to Forage and submit the task — learner-confirmed; completion page visible.

## Safety notes

- Keep the original source workbook unchanged.
- Use a local working copy for macro practice.
- Keep Excel's notification-based macro security. Do not enable all macros globally.
- Run only macros you have reviewed and understand.
- Do not upload the resulting `.xlsm` or republish the provided dataset unless redistribution is allowed by the provider.
