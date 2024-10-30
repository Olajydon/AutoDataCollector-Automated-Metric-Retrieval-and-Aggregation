# AutoDataCollector-Automated-Metric-Retrieval-and-Aggregation
AutoDataCollector is a VBA tool that automates the retrieval and aggregation of performance metrics across Excel workbooks. Triggered by a "Get Data" button, it dynamically fetches, filters, and summarizes data based on specified criteria, saving time and ensuring accuracy in reporting.

## Purpose
AutoDataCollector was initially created to streamline the collection and aggregation of key performance metrics across multiple Excel files. By automating the process, it saves time, reduces manual data entry errors, and ensures consistent, accurate reporting.

## Problem

In many data-driven environments, generating performance reports involves pulling data from different sources, filtering it based on specific criteria, and then aggregating it manually. This process can be time-consuming and error-prone, particularly when reports need to be produced regularly.

## Solution

The AutoDataCollector macro solves this problem by:
- Automatically retrieving data from a specified data source.
- Filtering data based on predefined metrics and other criteria.
- Aggregating data as required and populating it into a specified target range.

Triggered by a "Get Data" button, the macro is structured to allow for easy modification of source data locations, metric names, and aggregation criteria, making it adaptable to different reporting needs.

## Implementation Details

- **Code Structure**: The VBA macro is organized to allow flexibility in metric definitions and data source customization.
- **Metric Flexibility**: The code uses an array of generic metric names (`metric_1`, `metric_2`, etc.), which can be replaced or expanded to match specific reporting requirements.
- **Data Source**: The data source is referred to as `data_source`, a placeholder for any workbook and sheet containing the relevant data.
- **Aggregation Criteria**: The macro includes additional conditions, such as summing values that match specific criteria (e.g., entries containing certain "metrics" and "search_word").

## Usage Instructions

For detailed setup and usage instructions, refer to the [AutoDataCollector Usage Guide](https://github.com/Olajydon/AutoDataCollector-Automated-Metric-Retrieval-and-Aggregation/blob/main/docs/AutoDataCollector_usage_example.md)
.
