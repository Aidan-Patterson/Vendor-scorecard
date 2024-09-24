# Vendor-Scorecard Macros

This Excel workbook is designed to automate the process of grading vendors based on various performance metrics. It filters and processes data related to Purchase Orders (PO Data), Non-Conformance Reports (NCR Data), Rework Costs, and Costs of Poor Quality. The workbook uses a combination of VBA macros, buttons, and shapes to help companies streamline vendor evaluation and scorecard generation. These macros enable efficient filtering and manipulation of data across different sheets to produce clear insights into vendor performance.

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

## What This Workbook Does

This workbook automates the vendor grading process by using data from various company processes. Specifically, it handles:

- **Purchase Order (PO) Data**: Filters and processes PO information, including delivery performance (e.g., early or on-time deliveries), to evaluate vendor reliability.
- **Non-Conformance Reports (NCR Data)**: Identifies and filters NCR entries that reflect vendor quality issues and helps grade vendors based on these reports.
- **Rework Costs**: Tracks and processes data related to rework efforts and costs due to vendor errors or defects, which affects the vendor grading.
- **Cost of Poor Quality**: Analyzes the total cost incurred from poor quality products and processes from vendors, which contributes to the vendor score.

This automated analysis helps companies generate vendor scorecards to assess vendor performance on key metrics like delivery timeliness, product quality, and cost efficiency.

## Accessing the Macros via Buttons and Shapes

This workbook contains buttons and shapes on various sheets that allow users to run the macros directly without needing to navigate through the Developer tab or the VBA editor.

### Buttons and Shapes Overview

1. **Filter Data by Quarter**
   - **Button and Shape Location**: These are located on the **'Printout'** sheet.
   - **Functionality**: After you input the desired quarter in the **'Printout'** sheet (e.g., **Quarter 1**, **Quarter 2**), click the button or shape to filter the data across the relevant sheets and process it accordingly.

2. **Process Input Data**
   - **Button and Shape Location**: On the **'Input'** sheet, there are both buttons and shapes to process the input data.
   - **Functionality**: Clicking either the button or shape will process and update the data in other related sheets, based on the information provided in the **'Input'** sheet.

3. **Find Input Records**
   - **Button and Shape Location**: These are located on the **'Input Finder'** sheet.
   - **Functionality**: Use the button or shape to locate and process specific records from the **'Input'** sheet by filtering and finding the required data.

4. **Analyze Cost of Poor Quality**
   - **Button and Shape Location**: On the **'Cost of Poor Quality'** sheet.
   - **Functionality**: Clicking the button or shape triggers an analysis of the data related to the cost of poor quality and generates a report or output within the workbook.

## Additional Tools and References

1. **Missing References**:  
   If you encounter the "Can't find project or library" error, it means there is a missing reference in the VBA project. To resolve this:
   - Open the VBA editor (`Alt + F11`).
   - Go to **Tools** > **References**.
   - Uncheck any references marked as "MISSING".
   - Click **OK** to save.

2. **Resetting VBA Libraries**:  
   Ensure that common libraries like **Microsoft Office Object Library** and **Microsoft Excel Object Library** are selected in the References window.

## Issues and Troubleshooting

### Common Errors
- **Missing References**:  
  If a library reference is missing, follow the instructions under the **Additional Tools and References** section.
  
- **Button or Shape Not Working**:  
  If a button or shape doesn’t work when clicked, make sure macros are enabled (see the **Prerequisites** section). If macros are enabled but the button or shape still doesn’t work, check for errors in the VBA code via the editor (`Alt + F11`).

### Debugging Tips
- To step through the code, use the **F8** key in the VBA editor. This will allow you to run the code line by line and inspect variables for debugging.

## How to Edit the Macros

If you need to modify or review the macros:
1. Press `Alt + F11` to open the **Visual Basic for Applications (VBA)** editor.
2. In the Project Explorer, navigate through the workbook’s code modules.

## Summary

This Excel workbook simplifies data processing and vendor grading by using a combination of macros, buttons, and shape-driven workflows. The following sheets contain buttons and shapes to run specific macros:
- **Printout**: Filter data by quarter.
- **Input**: Process input data.
- **Input Finder**: Find and process input records.
- **Cost of Poor Quality**: Analyze and report on the cost of poor quality.

Ensure macros are enabled and that you have the required references set up. If you encounter issues, refer to the **Issues and Troubleshooting** section for guidance.
