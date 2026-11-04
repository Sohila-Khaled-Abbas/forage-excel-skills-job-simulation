# Task 04 — Sorting Macros and Buttons

**Status:** Complete — learner-confirmed. The learner reports 8/8 completion: the first recorded macro has a Form Control button; `SortBy5YrCAGR` and `SortBy2021UnitSales` were recorded; each sorting macro has its own button; both buttons passed their tests; and the two required screenshots were assembled into a PDF that was reopened and checked for readability. The macro-enabled working copy and PDF are both present in the repository. Forage upload and submission are still pending confirmation.

## 1. Continue in the same Task 04 workbook

Use the same local macro-enabled file, for example `Account_Sales_Data_for_Automation_Practice.xlsm`.

- Source workbook: [Account_Sales_Data_for_Automation.xlsx](../material/04-vba-and-marcos/Account_Sales_Data_for_Automation.xlsx)
- Macro-enabled working copy: [Account_Sales_Data_for_Automation_Practice.xlsm](../material/04-vba-and-marcos/Account_Sales_Data_for_Automation_Practice.xlsm)
- Completed evidence PDF: [Task_04_Macro_Buttons_Evidence.pdf](../material/04-vba-and-marcos/Task_04_Macro_Buttons_Evidence.pdf)
- Keep the original XLSX unchanged.
- Continue using desktop Excel.
- Save the XLSM before making changes.

Do not switch to the Task 03 dashboard workbook. This is a separate exercise and should retain its own macros.

## 2. Add a Form Control button for the first recorded macro

1. Select **Developer → Insert**.
2. In **Form Controls**, choose **Button** (not ActiveX Controls).
3. Click and drag to draw the button in a clear area of the worksheet.
4. In the **Assign Macro** dialog, select the macro you recorded in the preceding exercise.
5. Select **OK**.
6. Right-click the button → **Edit Text**.
7. Give it a concise label that describes the actual action, for example **Format Headers** if that was the action you recorded.
8. Click the button once to verify that it runs the correct macro.

If the workbook data fills columns A:M, place buttons in an unused area to the right of the dataset, for example columns O:R. Do not cover any headers, records, or data cells. If the sheet uses more columns, choose another genuinely blank area.

If the wrong macro is assigned, right-click the button and select **Assign Macro** to correct it.

Microsoft's official instructions: [Assign a macro to a Form or a Control button](https://support.microsoft.com/en-us/excel/assign-a-macro-to-a-form-or-a-control-button).

## 3. Confirm the complete data range before sorting

The previous Account Sales dataset layout used `A1:M41` (one header row and 40 account rows), with:

- `2021 Unit Sales` in column L.
- `5 YR CAGR` in column M.

**Verify these in the Task 04 workbook rather than assuming the same layout.** Check the actual worksheet tab, header row, first record, last record, and last data column. If the Task 04 workbook matches the prior layout, `A1:M41` is the complete range. Otherwise use the actual contiguous range.

Before recording either sort:

1. Make sure no rows or columns in the data range are hidden.
2. Confirm all account fields are in one contiguous rectangular range.
3. Confirm that the first row contains the column labels.
4. Use the Name Box (left of the formula bar) to select the complete range, including its header, while recording. For example, type `A1:M41` only if that range has been confirmed.
5. In the Sort dialog, keep **My data has headers** checked.

Do not sort only column L or only column M. Sorting one field alone can detach values from their corresponding account records. Excel's sort dialog supports sorting a data range by one selected column while moving complete rows together: [Sort data in a range or table in Excel](https://support.microsoft.com/en-us/excel/sort-data-in-a-range-or-table-in-excel).

## 4. Record macro 1 — Sort by 5 YR CAGR

Macro name (no spaces):

```text
SortBy5YrCAGR
```

Recording sequence:

1. Select **Developer → Record Macro**.
2. Enter `SortBy5YrCAGR`.
3. Set **Store macro in** to **This Workbook**.
4. Select **OK** to start recording.
5. Select the entire verified data range using the Name Box, including the header row.
6. Open **Data → Sort**.
7. Make sure **My data has headers** is checked.
8. Set:
   - **Sort by:** `5 YR CAGR`
   - **Sort On:** Cell Values
   - **Order:** Largest to Smallest
9. Select **OK**.
10. If Excel displays a Sort Warning, choose **Expand the selection** so all columns in each account row move together.
11. Return to **Developer → Stop Recording**.

The sheet should now be ordered by the CAGR field. Do not type replacement values into the data.

## 5. Record macro 2 — Sort by 2021 Unit Sales

Macro name:

```text
SortBy2021UnitSales
```

Repeat the same recording workflow, using a fresh **Developer → Record Macro** session and the full data range.

In the Sort dialog set:

- **Sort by:** `2021 Unit Sales`
- **Sort On:** Cell Values
- **Order:** Largest to Smallest
- **My data has headers:** checked

Stop recording when the sort completes.

## 6. Add one button for each sorting macro

Create two more **Form Control → Button** controls.

| Button label | Assign this macro |
|---|---|
| Sort by 5 YR CAGR | `SortBy5YrCAGR` |
| Sort by 2021 Unit Sales | `SortBy2021UnitSales` |

For each button:

1. Developer → Insert → Button (Form Control).
2. Draw the button in the free area beside the dataset.
3. Select the matching macro in **Assign Macro**.
4. Right-click → **Edit Text** and set the label shown above.
5. Resize the button so the full label is readable.

Keep both buttons and any earlier macro button visible together for the evidence screenshot. Do not position buttons over the data.

## 7. Test both sort buttons

Test one button at a time:

### Test A — CAGR

1. Click **Sort by 5 YR CAGR**.
2. Confirm the 5 YR CAGR values are in descending order (highest numeric value first).
3. Confirm the header row stays at the top.
4. Check that account names and the other columns remain paired across each row.

### Test B — 2021 Unit Sales

1. Click **Sort by 2021 Unit Sales**.
2. Confirm the 2021 Unit Sales values are in descending order (largest first).
3. Confirm the header row stays at the top.
4. Check that the complete account rows move together.

Each button should run its own macro. If a button runs the wrong sort, right-click it → **Assign Macro**, select the correct procedure, and test again.

**Data-quality caution:** If the CAGR column contains text placeholders such as `Formula Error`, do not assume they are numeric CAGR values. Mixed text and numbers have separate sort behavior; blank cells remain last in Excel's documented sort order. Inspect the result and keep the original markers intact. Do not overwrite source values merely to force the ranking. Microsoft reference: [Sort data in a range or table in Excel](https://support.microsoft.com/en-us/excel/sort-data-in-a-range-or-table-in-excel).

## 8. Prepare the PDF submission

Forage asks for a PDF containing two pieces of evidence:

1. A screenshot of the workbook showing both sorting macro buttons.
2. A screenshot of the VBA code for one sorting macro.

Suggested process:

1. In Excel, set a zoom that shows the full data area and both buttons clearly.
2. Capture the worksheet with **Win + Shift + S** (Snipping Tool). Make sure the two sort buttons are visible and not covered by menus.
3. Press **Alt + F11** to open the Visual Basic Editor.
4. In Project Explorer, open the standard module containing `SortBy5YrCAGR` or `SortBy2021UnitSales`.
5. Scroll so the complete procedure is readable, then capture a second screenshot.
6. Paste both screenshots into a Word document, add short captions, and arrange them on one or two pages.
7. Select **File → Save As** or **Export → Create PDF/XPS** and save as `Task_04_Macro_Buttons_Evidence.pdf`.
8. Reopen the PDF and check that both screenshots are legible and no code is clipped.

Do not include login information, personal files, or unrelated desktop content in screenshots. The PDF should contain evidence, not the supplied source workbook.

## 9. Final completion checklist

- [x] Form Control button added for the first recorded macro.
- [x] `SortBy5YrCAGR` recorded.
- [x] `SortBy2021UnitSales` recorded.
- [x] Each sorting macro assigned to its own Form Control button.
- [x] CAGR sorting button tested successfully.
- [x] 2021 Unit Sales sorting button tested successfully.
- [x] Workbook screenshot showing both sorting buttons captured.
- [x] VBA code screenshot captured.
- [x] Both screenshots assembled into `Task_04_Macro_Buttons_Evidence.pdf`.
- [x] PDF reopened and checked for readability.
- [x] Macro-enabled working copy and evidence PDF present in the repository; original XLSX retained.
- [ ] Upload the macro-enabled XLSM and PDF to the Forage task page and select **Submit file**.

**Completion boundary:** The Excel implementation and evidence preparation are complete on a learner-confirmed basis. Do not mark the Forage activity submitted until the upload is accepted and the platform confirms the submission.

## 10. Safety and reproducibility

- Record each macro in the Task 04 XLSM using the Macro Recorder as required by the learning exercise.
- Do not paste pre-written macros in place of the recording exercise.
- Keep each macro limited to one sorting operation.
- Keep the complete table/range together when sorting; never sort a single data column by itself.
- Save the workbook after successful tests.
- Keep the macro-enabled workbook local unless the provider's terms allow uploading it. The repository may contain the learning guide and evidence screenshots only if sharing permissions allow.
