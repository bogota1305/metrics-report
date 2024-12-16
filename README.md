# Business Metrics Extraction and Reporting Tool

## Overview

This project provides a set of Python scripts designed to extract, process, and visualize business metrics from a database. The primary goal is to generate comprehensive Excel reports that track key performance indicators (KPIs) on a daily basis.

## Features

The project currently supports the extraction and reporting of three main types of metrics:

1. **Order Metrics**
2. **Sales and Renewals Metrics**
3. **Payment Error Metrics**

### Key Functionalities

- Retrieve data from a MySQL database for a specified date range
- Process and aggregate data on a daily basis
- Generate line charts and bar charts for visual analysis
- Export results to Excel spreadsheets with multiple sheets and visualizations

## Project Structure

### Main Scripts

1. `orders.py`

This script focuses on querying and processing sales data from the database, differentiating between new and old users, as well as recurring and non-recurring orders. The results are visualised in graphs embedded in an Excel file.
  
  - Functions:
  
    - Data Query
    Executes SQL queries to retrieve order data from bi.fact_orders and bi.fact_sales_order_items.
    Filters data by dates and excludes statuses such as ‘CANCELLED’ and ‘PAYMENT_ERROR’.
    
    - Data Processing
    Split data into subgroups: new users (SUBS, OTO, MIX) and old users (SUBS, OTO, MIX), plus recurring orders.
    Adds additional columns such as date and performs daily calculations.
         * Total revenue
         * Number of sales
         * Total items sold
         * Average items per order
         * Average order value
    
    - Charts and Excel
    Generate charts to visualise metrics such as revenue, number of sales and average per order.
    Use the save_dataframe_to_excel module to save this data to an Excel file with charts.


2. `payments.py`
   
This script processes payment error data from the database, segmented by groups and dates. The results are analysed to identify patterns of errors and graphs and reports are generated in an Excel file.

  - Functions:
  
    - Data Query
    Execute a SQL query to get payment data from sales_and_subscriptions.payments.
    Filter by date using query_start_date, calculated as one month before start_date.
    
    - Data Processing
    Group data by entityId and calculate metrics related to errors and resolutions.
    Identifies the first error and the date the error was successfully resolved.
    
    - Metrics Calculation
    Creates daily metrics for errors and resolutions.
    Calculates overall totals for errors and resolutions.
    
    - Error Reason Analysis
    Extract declination codes based on error metadata.
    
    - Charting and Reporting
    Generates graphs to visualise daily errors and resolutions.
    Saves data to an Excel file using save_dataframe_to_excel and save_error_reasons_with_chart.


3. `renewalsAndNoRecurrents.py`
   
This script is responsible for processing sales data from the database, performing daily calculations and plotting the results in an Excel file.   

  - Functions:
  
    - Data Query
    Executes an SQL query to retrieve data from bi.fact_orders between the specified dates (start_date and end_date).
    Excludes cancelled and payment error orders.
    
    - Data Processing
    Groups data by date and separates recurring and non-recurring orders.
    Calculates daily metrics and grand totals.
    Charts
    
    - Excel generation
    Uses the save_dataframe_to_excel module to save processed data to an Excel file with embedded charts.


4. `main.py`
   
This file serves as an entry point for the management and execution of processes related to orders, sales and payment errors based on user-selected dates.

### Supporting Modules

- `date_selector.py`:
  
This script uses the Tkinter library to create a graphical interface that allows users to select dates and report options to generate an Excel file with specific data.

  - Functions:

    - Graphical interface
    Allows the user to select a start date and an end date using an interactive calendar (tkcalendar).
    Users can enter a name for the file in which the processed data will be saved.
    
    - Variable Selection
    Provides options to select different types of reports, including all orders, single orders, sales and payment errors.
    Provides quick select/deselect buttons and options to enable/disable specific sections of reports.
    
    - Report Generation
    Once dates and options are selected, the script returns these values along with the desired file name and options selected.

    - Data Validation
    Verifies that all required fields are complete before proceeding with report generation.


- `excel_creator.py`:

This script performs Excel file generation using the openpyxl library, as well as creating graphs using matplotlib. The data is processed based on the selection of variables and metrics, and stored in Excel sheets with relevant graphs.

  - Functions:
  
    - save_dataframe_to_excel
    Creates an .xlsx file with different sheets including tabular data and charts.
    Uses matplotlib to generate charts as images and integrates them into the Excel sheets.
    
    - line_chart
    Creates a line chart to visualise selected metrics per day and saves it as an image, then inserts it into the Excel sheet.
    
    - bar_chart
    Generates a bar chart to compare metrics between different data sets, such as new and old, recurring and non-recurring.
    
    - save_error_reasons_with_chart
    Saves error reasons in a new Excel sheet with dynamic colours based on error types.
    Adds a bar chart showing the number of errors by type.
    
    - save_dataframe_to_excel_orders
    Creates a sheet with graphs and processed data related to specific orders.


- `database_queries.py`:

The execute_query function allows you to execute an SQL query and return the results in a pandas DataFrame. This makes it easy to work with data extracted from a MySQL database and manipulate or analyse it using data analysis operations.


- `colors.py`:

The lighten_color function takes a hexadecimal colour as input and returns a lighter colour based on an intensity factor (default to 0.5). This is useful for adjusting colours in charts or spreadsheets, providing smoother visual variety.

## Prerequisites

- Python 3.8+
- Required Libraries:
  * pandas
  * mysql-connector-python
  * matplotlib
  * openpyxl
  * tkinter
  * tkcalendar

## Installation

1. Clone the repository
2. Install required dependencies:
   ```bash
   pip install pandas mysql-connector-python matplotlib openpyxl tkinter tkcalendar
   ```

## Usage

1. Modify the database_queries.py file inside the modules folder and assign the corresponding values to the host, user and password variables to connect to the database.
2. Run the file main.py
3. Use the GUI to:
   - Select start and end dates. Choose the end date one day after the day you want (e.g. if you want the data from 1 to 10 October, choose 1 October as the start date and 11 October as the end date).
   - Choose report types
   - Specify output folder name

Example report types:
- New user orders (SUBSCRIPTION/OTO/MIXED)
- Existing user orders
- Recurrent orders
- Sales breakdown
- Payment error analysis

## Configuration

- Ensure database connection details are correctly configured in the database connection module
- Modify SQL queries in respective scripts if database schema changes

## Output

Each report generates an Excel file with:
- Daily metrics table
- Line charts for key metrics
- Optional comparative bar charts
- Detailed breakdowns by user type and plan

