VBAFinancialAnalysis
====================

Excel VBA code to clean, represent financial timeseries in a plot, list and anaklysis output:

1. This script prompts VBA to locate data.csv file which is stored locally and contains two rows of data: 

- timestamp
- price data

2. This script will clean the data and will try to fill in gaps where price data does not have specific values 
   as interpolation of simple average.

3. It will create a new workbook and make sure 3 sheets to write to are also created

4. Graph will be created from a 2d matrix data array in VBA

5. Timeseries listing will be printed with price data quared on the next sheet.

6. Average timeseries price points will be posted into the average sheet.

PS: you can download sample data.csv from this repository and then refer to it from the prompt in the VBA.

Enjoy!
Vasyl
