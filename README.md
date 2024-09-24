# vendor-scorecard

This repository contains an Excel workbook with VBA macros designed to automate vendor scorecard processing. The macros perform tasks such as filtering data, manipulating specific worksheets, and processing information based on user inputs.

## Prerequisites

Before using the macros in this workbook, ensure that:
- You are using **Microsoft Excel**.
- You have enabled **macros** and **VBA** in Excel.
- You have any necessary libraries installed (described below).

### How to Enable Macros in Excel
To use the macros in this workbook:
1. Open Excel and go to **File** > **Options**.
2. Select **Trust Center** > **Trust Center Settings**.
3. Choose **Macro Settings**.
4. Select **Enable all macros** (Note: This might pose security risks, so only enable macros from trusted sources).
5. Click **OK** to save your settings.

## Accessing the Macros

To view or edit the macros:
1. Press `Alt + F11` to open the **Visual Basic for Applications (VBA)** editor.
2. In the Project Explorer, navigate through the workbook’s code modules.

## Macros Overview

### 1. `FilterDataByQuarter`
This macro filters data based on the financial quarter provided in the **'Input'** sheet. It applies filters to the following worksheets:
- **NCR Data**
- **Rework Data**
- **Response Data**

After filtering the data based on the specified quarter, the macro processes it and outputs it as required.

### 2. `ProcessCompanyData`
This macro aggregates data for vendors listed in the **'PO Data'** sheet and outputs the result to the **'PO DataOutput'** sheet. It sums the relevant performance metrics for each vendor, including **'Early'** and **'On-Time'** performance.

### 3. `EnterReworkData`
This macro processes rework data by taking inputs from the **'Input'** sheet and updating the corresponding rows in the **'Rework Data'** sheet. The same data is also copied to an external workbook for further analysis.

## Additional Tools and References

1. **Missing References:**
   If you encounter the "Can't find project or library" error, it means there is a missing reference in the VBA project. To resolve this:
   - Open the VBA editor (`Alt + F11`).
   - Go to **Tools** > **References**.
   - Uncheck any references marked as "MISSING".
   - Click **OK** to save.

2. **Resetting VBA Libraries:**
   - Ensure that common libraries like **Microsoft Office Object Library** and **Microsoft Excel Object Library** are selected in the References window.

3. **External Workbook Reference:**
   If your workbook interacts with another workbook (e.g., **Vendor Scorecard TEST.xlsm**), ensure that workbook is available in the specified directory. You may need to adjust file paths within the macro code if your directory structure differs.

## Using the Macros

1. **Filter Data by Quarter:**
   - Enter the desired quarter in the **'Input'** sheet (e.g., **Quarter 1**, **Quarter 2**).
   - Run the `FilterDataByQuarter` macro by going to the **Developer Tab** and selecting **Macros**, then click **Run**.

2. **Process Vendor Data:**
   - Ensure all relevant vendor data is in the **'PO Data'** sheet.
   - Run the `ProcessCompanyData` macro to aggregate and summarize the data in **'PO DataOutput'**.

3. **Enter Rework Data:**
   - Ensure the necessary data is entered in the **'Input'** sheet.
   - Run the `EnterReworkData` macro to update both the **'Rework Data'** sheet in the current workbook and in the external **Vendor Scorecard TEST.xlsm** workbook.

## Issues and Troubleshooting

### Common Errors
- **Missing References:** If a library reference is missing, follow the instructions under the **Additional Tools and References** section.
- **Macro Not Running:** Ensure macros are enabled, as detailed in the **Prerequisites** section. If the macro still fails, check for errors in the VBA code via the editor (`Alt + F11`).

### Debugging Tips
- To step through the code, use the **F8** key in the VBA editor. This will allow you to run the code line by line and inspect variables for debugging.

