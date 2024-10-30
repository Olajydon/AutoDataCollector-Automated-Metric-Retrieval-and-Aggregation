# Excel Automation Project: Data Collection and Visualization

This repository contains VBA tools for automating data collection, aggregation, and visualization across Excel workbooks. The **AutoDataCollector** and **RefreshChart** macros work together to streamline data retrieval and chart updates, making performance reporting efficient and accurate.

## AutoDataCollector: Automated Metric Retrieval and Aggregation

**Description**:  
The **AutoDataCollector** macro automates the retrieval and aggregation of performance metrics across Excel workbooks. Triggered by a "Get Data" button, it dynamically fetches, filters, and summarizes data based on specified criteria, saving time and ensuring reporting accuracy.

### Purpose

AutoDataCollector was created to simplify the process of collecting and aggregating key performance metrics from multiple Excel files, reducing manual work and ensuring consistency.

### Problem

In data-driven environments, generating performance reports often involves gathering data from various sources, filtering it, and manually aggregating it. This process is time-intensive and error-prone, especially for regular reporting.

### Solution

The AutoDataCollector macro solves this problem by:
- Automatically retrieving data from a specified source.
- Filtering data based on predefined metrics.
- Aggregating and populating data into a designated target range.

The macro is designed for easy modification of source data locations, metric names, and aggregation criteria, making it adaptable to a wide range of reporting needs.

### Implementation Details

- **Code Structure**: The macro allows flexibility in defining metrics and customizing data sources.
- **Metric Flexibility**: Generic metric names (`metric_1`, `metric_2`, etc.) allow adaptation for different reporting needs.
- **Data Source**: Uses a placeholder, `data_source`, for any workbook or sheet containing relevant data.
- **Aggregation Criteria**: The code includes optional conditions for summing values that match specific criteria.

### Usage Instructions

For detailed setup and usage instructions, refer to the [AutoDataCollector Usage Guide](docs/AutoDataCollector_usage_example.md).

---

## RefreshChart: Automated Chart Refresh for Performance Tracking

**Description**:  
The **RefreshChart** macro dynamically updates charts to display the latest data, ensuring real-time performance tracking without manual adjustments. Designed for flexibility, it supports updates across multiple metrics and is easily adaptable for other reporting scenarios.

### Purpose

The RefreshChart macro ensures that charts always reflect the most current data, which is essential for accurate performance tracking.

### Problem

As data updates frequently in performance reports, manually refreshing charts to reflect this data can be both time-consuming and prone to error, especially when multiple metrics are involved.

### Solution

The RefreshChart macro addresses this by:
- Automatically updating charts based on the latest data in specified ranges.
- Ensuring all relevant metrics are consistently refreshed.

This tool streamlines the reporting process by enabling quick, reliable chart updates with minimal manual intervention.

### Implementation Details

- **Chart Positioning**: Places the chart in a predefined location based on the sheet type.
- **Dynamic Data Range**: The macro identifies the latest filled data and updates the chart accordingly.
- **Multi-Sheet Compatibility**: Supports updates across different sheet types, with specific handling for “special” sheets.

### Usage Instructions

For detailed setup and usage instructions, refer to the [RefreshChart Usage Guide](docs/RefreshChart_usage_example.md).
