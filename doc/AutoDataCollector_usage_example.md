# AutoDataCollector Usage Guide

## Purpose
The **AutoDataCollector** macro automates data retrieval and aggregation across Excel workbooks, enabling efficient and accurate reporting.

## Instructions

1. **Add a "Get Data" Button**:
   - Go to the **Developer** tab in Excel (if the Developer tab isn’t visible, enable it in Excel’s options under **Customize Ribbon**).
   - Click **Insert** > **Button (Form Control)**, and place it on the sheet where you want the data to populate.
   - When prompted, assign the `AutoDataCollector` macro to this button.
   - Edit the button text to display "Get Data" or aany suitable name.

2. **Select Target Cells**:
   - In the active sheet, highlight the cells where each metric’s data should appear.

3. **Run the Macro**:
   - Click the "Get Data" button to start the AutoDataCollector macro.
   - The macro will prompt you to select a source workbook if it isn’t already open.
   - It will retrieve, filter, and aggregate the data for each metric based on the visible rows in the source sheet and populate the selected cells.

## Adaptability
- **Customizing Metrics**: Modify the `metrics` array in the code (e.g., `metric_1`, `metric_2`, etc.) to match the metrics relevant to your reporting needs.
- **Adjusting the Data Source**: Update the `data_source` reference in the code to point to any workbook and sheet where your data is stored.

