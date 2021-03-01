# Stock Analysis: Refactoring VBA Macro & Excel Spreadsheet Visualization

## Overview of Project
*Conducting Financial Analysis on Multiple Stock Tickers Through VBA Macro Automation in Excel*
---
### Purpose
The following project involves a deliberate, strategic refactoring of the Stock Market dataset, which utilzes VBA code scripting. The code solution loops through the entirity of the dataset, supported by conditional logic functions (if-then statements), in order to calculate KPI variables in the market that measure the perfromance of a tickerIndex over a given period of time.  These metrics of performance are treated as arrays within the VBA code: tickerVolumes, startingTickerPrices and endingTickerPRices. In this analysis specifically, we are looking to find the annual Return, defined as the variance in stock price within a given year. The exercise of refactoring, or editing the code, is to demontrate the advantages of making the script more efficient--this is done by taking fewer steps to loop through all the data to find the output variables. Conducting a refactoring should in turn make the macro script more reader friendly and use less computer memory; therefore, it should theoretically take Excel a shorter amount of time to run the analysis.

---
### Results
The following delinates the results of "AllStocksAnalysisRefactored" Subroutine, which can be used to compare the stock performance in tickerIndex, between 2017 and 2018:

- 1a) Final VBA Refactoring: 2017 Analysis

  - ![VBA_Challenge_2017_outputs](C:\Users\carly\OneDrive\Desktop\data_bootcamp\analysis_projects\green_stocks_analysisVBA\Challenge Deliverables\RESOURCES\VBA_Challenge_2017._outputs.png)

  - ![VBA_Challenge_2017](C:\Users\carly\OneDrive\Desktop\data_bootcamp\analysis_projects\green_stocks_analysisVBA\Challenge Deliverables\RESOURCES\VBA_Challenge_2017.png)

- 1b) Final VBA Refactoring: 2018 Analysis

  - ![VBA_Challenge_2018_outputs](C:\Users\carly\OneDrive\Desktop\data_bootcamp\analysis_projects\green_stocks_analysisVBA\Challenge Deliverables\RESOURCES\VBA_Challenge_2018._outputs.png)

  - ![VBA_Challenge_2018](C:\Users\carly\OneDrive\Desktop\data_bootcamp\analysis_projects\green_stocks_analysisVBA\Challenge Deliverables\RESOURCES\VBA_Challenge_2018.png)

Below is the comparison between the execution times of the original VBA script and the refactored VBA script:

- 2a) Original Script Execution Time 

  - ![OriginalScript_ExecutionTime](C:\Users\carly\OneDrive\Desktop\data_bootcamp\analysis_projects\green_stocks_analysisVBA\Challenge Deliverables\RESOURCES\OriginalScript_ExecutionTime.png)

- 2b) Refactored Script Execution Time

  - ![RefactoredScript_ExecutionTime](C:\Users\carly\OneDrive\Desktop\data_bootcamp\analysis_projects\green_stocks_analysisVBA\Challenge Deliverables\RESOURCES\RefactoredScript_ExecutionTime.png)

---
### Summary of Post-Refactoring Reflection

#### Advantages & Disadvantages of Refactoring Code

_Advantages_
- Lowers machine processing time
- Automates analysis with lower frequency of ad-hoc updating (time-saving)
- More user-friendly and readable to coder's & following editor's eyes
- Aids in the process of finding sloppy/pseudocode & potential error bugging

_Diadvantages_
- The benefits of time saving and machine processing depend on the project at hand
  - Must weigh the cost of refactoring (time and money) to benefit of code output performance
- If project is large, there is potential to crash code in totality with steep debugging
- Problematic when code is sent off to different programmer who doesn't comprehend code

Overall, refactoring is one of a programmers most valuable assets, as it is utilized to update and maintain code script over time as dataset grows and evolves in analytical metrics. Intuitively, refactoring is a necessary step of the analysis refresh when a dataset requires updating. As evident by the execution time images above, refactoring also provides a programmer with improved functionality of code, allowing for faster processing times and cleaner script. And when thought through properly from the get, refactoring can save a programmer a lot of time and manual hastle in the long run.
