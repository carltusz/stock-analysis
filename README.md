# stock-analysis
Carleton Analytics Bootcamp Module 2: Independent Learning

## Overview
This is a VBA project to analyze stock data on behalf of a hypothetical client. Data has been gathered for green industry stocks from the years 2017 and 2018 for analysis. A script was developed in Excel to analyse this data and provide the Total Daily Volume and Return for each stock.

## Results
### 2017
Green stocks in 2017 generally offered a positive return, with the exception of TERP, which lost 7.2% of value over the year. The greates gains were with DQ, SEDG, and ENPH, at 199.4%, 184.5%, and 129.5%, respectively. 

![2017 Results](./Resources/VBA_Challenge_2017_Results.png)

### 2018
Green stocks in 2018 generally offered a negative return, with the exception of RUN and ENPH which gained 84.0% and 81.9%, respectively.

![2018 Results](./Resources/VBA_Challenge_2018_Results.png)

### Script Performance
The data was processed using a VBA script. Using the original script, the queries were performed in approximately 0.16 seconds for both 2017 and 2018 datasets.

![2017 Original Runtime](./Resources/VBA_Challenge_2017_ClassworkRuntime.png) ![2018 Original Runtime](./Resources/VBA_Challenge_2018_ClassworkRuntime.png) 

Refactoring this script improved runtime significantly, making it closer to 0.06 seconds for either dataset. 

![2017 Refactored Runtime](./Resources/VBA_Challenge_2017.png) 
![2018 Refactored Runtime](./Resources/VBA_Challenge_2018.png) 

Both sets of data contained 3012 rows of raw data which were processed, and 12 tickers which were summarized.

## Summary
Refactoring code is an exercise wherin existing code is updated to improve performance, readability, or other attributes. The advantage is the presumed efficiency increase or maintainability of the new version of the code; however, refactoring itself can take a significant amount of time. Depending on the state of the existing code, refactoring it may not have significant benefit.

In the case of this report, the refactored code ran in less than half the time of the original code. While these relative time differences are significant, the absolute time savings is fairly insiginificant. The time spent refactoring the code was many time that of the actual runtime difference between the two versions. Given the relatively small, and finite, data set used here, refactoring the code did not offer any real benefits to the end user. If they continue to use a similarly formatted raw data set in years to come, the time savings would not add up to the time spent refactoring.

Refactoring would have a greater benefit if many more tickers were being tracked. The original code did recursive loops, meaning that it looped through each ticker on each line of raw data. The recursive looping means that the logic is run up to 11 times on each row of data, or close to 33,000 times per run. The refactored code ran it only once, or around 3000 times per run. This difference would become much greater with increasing dataset size.

If a larger set of data is used, it would also be useful to make the script more robust. The current architecture is highly dependent on "Magic Numbers" which are hard-coded in to the script.  A larger dataset, or varied dataset would likely justify more robust code to support varied formats and data organization. 

