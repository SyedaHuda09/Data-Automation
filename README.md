This project is designed to process Excel files containing financial data, particularly related to floor and interest rate values in investment records. It merges the values from Rate(b) and Floor(b) columns into a new column Interest rate with proper formatting. The tool can handle complex data scenarios such as missing values and different data types using regex patterns and flexible error handling.

***Features***1
1. Regex-Based Extraction: Uses regular expressions to extract and format percentage-based floor values from the Floor(b) column.

2. 3Interest Rate Merging: Merges data from the Rate(b) and Floor(b) columns into a single formatted string for easier readability.

3. Error Handling: Utilizes try-except blocks to ensure smooth processing of data, even when encountering non-standard formats or unexpected values.

4.Excel File Processing: Reads financial data from Excel files, processes it, and outputs a new Excel file with the merged Interest rate column.
