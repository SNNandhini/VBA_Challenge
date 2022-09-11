# VBA_Challenge
VBA Homework: The VBA of Wall Street

## Multiple Year Stock Data Analysis
The stock data provided to analyse spans across 3 years 2018, 2019 and 2020 respectively. The raw data contains Ticker, Date, Opening Price, Highest Price, Lowest Price, Closing Price and Volume for each year. The list is sorted by Ticker and Date.

### VBA Script and Results
1. The VBA Script developed for this assignment can be found in the file (Multiple_year_stock_data.vbs)

2. The file (Multiple_year_stock_data.xlsm) contains the data along with the results and VBA code. 

3. The VBA Script does the following:

    -  The Data for each year under each worksheet is read through to create a summary (Ticker, Yearly Change, Percent Change and Total Stock Volume).
        
    -  Ticker (Column I): A unique list of Tickers from the data provided in Column A.

    -  Yearly Change (Column J): In the code for each Ticker, this is calculated by subtracting the first opening price from the last closing price. As the data provided is already sorted based on Date, no extra check is required to identify the first and last dates. 

    -  Percent Change (Column K): The Yearly Change above is divided by opening price and the percentage is calculated.

    -  Total Stock Volume (Column L): This is the sum of the Volume (column G) for each Ticker.

    -  The calculations above are performed on all the worksheets (2018, 2019 and 2020).

    -  The result for each year is captured as images in the Images_Stock_Data_Analysis.docx.

### Bonus
1. The VBA script above is updated to identify the Tickers with Greatest % Increase, Greatest % Decrease and Greatest Total Volume for each year:

    -  Greatest % Increase: The Ticker with the highest Percent Change value (Column K) is identified and printed along with the value in columns P and Q.

    -  Greatest % Decrease: The Ticker with the lowest Percent Change value (Column K) is identified and printed along with the value in columns P and Q.

    -  Greatest Total Volume: The Ticker with the highest Total Stock Volume value (Column K) is identified and printed along with the value in columns P and Q.

    -  The result for each year is captured as images in the Images_Stock_Data_Analysis.docx.

### Files Submitted
1. Multiple_year_stock_data.vbs

2. Multiple_year_stock_data.xlsm

3. Images_Stock_Data_Analysis.docx

### References
1. Champs, E., 2022. Excel Champs. [Online] 
Available at: https://excelchamps.com/vba/autofit/#AutoFit_Entire_Worksheet
[Accessed 09 September 2022].

2. Champs, E., 2022. Excel Champs. [Online] 
Available at: https://excelchamps.com/vba/functions/formatpercent/
[Accessed 08 September 2022].

3. Siddique, A. A., 2022. exceldemy. [Online] 
Available at: https://www.exceldemy.com/vba-format-currency-two-decimal-places/
[Accessed 08 September 2022].