# Vendor-Scorecard 

This repository contains an Excel workbook with VBA macros designed to automate vendor scorecard processing. The macros perform tasks such as filtering data, manipulating specific worksheets, and processing information based on user inputs. There are buttons on the Excel sheet to make using these macros easier for users.

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

## Accessing the Macros via Buttons

To make things easier, the workbook contains buttons on specific sheets that allow users to run the macros directly with a single click, without needing to navigate through the Developer tab or the VBA editor.

### Buttons Overview

1. **Filter Data by Quarter**  
   - **Button Location**: This button is located on the **'Input'** sheet.
   - **Functionality**: After you input the desired quarter (e.g., **Quarter 1**, **Quarter 2**, etc.), click the button to filter the data across the following sheets:
     - **NCR Data**
     - **Rework Data**
     - **Response Data**

2. **Process Vendor Data**  
   - **Button Location**: On the **'PO Data'** sheet, you will find a button for processing vendor data.
   - **Functionality**: Clicking this button will aggregate vendor performance data and output it into the **'PO DataOutput'** sheet. It will sum values for each vendor, including performance metrics such as **'Early'** and **'On-Time'** deliveries.

3. **Enter Rework Data**  
   - **Button Location**: On the **'Input'** sheet.
   - **Functionality**: After entering necessary data on the 'Input' sheet, click this button to automatically update the relevant rows in the **'Rework Data'** sheet. It will also update the data in an external workbook named **Vendor Scorecard TEST.xlsm**.

## Additional Tools and References

1. **Missing References**:  
   If you encounter the "Can't find project or library" error, it means there is a missing reference in the VBA project. To resolve this:
   - Open the VBA editor (`Alt + F11`).
   - Go to **Tools** > **References**.
   - Uncheck any references marked as "MISSING".
   - Click **OK** to save.

2. **Resetting VBA Libraries**:  
   Ensure that common libraries like **Microsoft Office Object Library** and **Microsoft Excel Object Library** are selected in the References window.

3. **External Workbook Reference**:  
   If your workbook interacts with another workbook (e.g., **Vendor Scorecard TEST.xlsm**), ensure that workbook is available in the specified directory. You may need to adjust file paths within the macro code if your directory structure differs.

## Issues and Troubleshooting

### Common Errors
- **Missing References**:  
  If a library reference is missing, follow the instructions under the **Additional Tools and References** section.
  
- **Button Not Working**:  
  If a button doesn’t work when clicked, make sure macros are enabled (see the **Prerequisites** section). If macros are enabled but the button still doesn’t work, check for errors in the VBA code via the editor (`Alt + F11`).

### Debugging Tips
- To step through the code, use the **F8** key in the VBA editor. This will allow you to run the code line by line and inspect variables for debugging.

## How to Edit the Macros

If you need to modify or review the macros:
1. Press `Alt + F11` to open the **Visual Basic for Applications (VBA)** editor.
2. In the Project Explorer, navigate through the workbook’s code modules.

## Summary

This Excel workbook simplifies vendor scorecard processing by using a combination of macros and button-driven workflows. You can:
- Filter data by quarter using the **Filter Data** button on the **'Input'** sheet.
- Process vendor data using the **Process Vendor Data** button on the **'PO Data'** sheet.
- Enter rework data using the **Enter Rework Data** button on the **'Input'** sheet.

Ensure macros are enabled and that you have the required references set up. If you encounter issues, refer to the **Issues and Troubleshooting** section for guidance.

