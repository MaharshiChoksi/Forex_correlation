# Forex Correlation chart

## Project Overview
This project utilizes the MT5 API to retrieve historical data for analyzing currencies and indices.

Here’s a structured outline of the development process:

## OVERVIEW
### 1. Connect to MT5
* Establish a connection to the MT5 terminal and authenticate using user credentials.
### 2. Retrieve Historical Data
* Pull historical data for specified tickers (currency pairs and indices) from the MT5 server using the API. 
* Data is retrieved for various timeframes relative to a specified starting time.
### 3. Data Handling and storage
* Store the retrieved data in a Pandas DataFrame.
* Replace any null values with zeros to ensure data consistency.
* Export the cleaned data to an Excel file for further analysis.
### 4. Calculate Correlation
* Implement a separate program (calculate_correlation) to analyze the data:
  * Parse the Excel file from the argument.
  * Read each worksheet into a Pandas array for processing.
  * Utilize the pearsonr method from the stats module in scipy to compute correlation coefficients between different currency pairs or indices.
  * Store the correlation coefficient data back into respective sheets within the Excel file.
### 5. Visual Representation
* Apply conditional formatting to the Excel sheets to visually depict correlation strengths:
  * Cells are colored differently based on correlation values, facilitating easier interpretation of relationships.