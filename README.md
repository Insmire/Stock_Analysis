<!-- Overview of Project: Explain the purpose of this analysis. -->
# Overview
The purpose of this project is to refactor the AllStocksAnalysis subroutine in MS Excel Visual Basics for Application (VBA) with subroutine AllStocksAnalysisRefactored in order to execute 2017 and 2018 stock analysis more efficientely. Same data is utilized in AllStocksAnalysis as in AllStocksAnalysisRefactored. Due to the nature of VBA running all past saved codes regardless of any present editions, all refactored code is created and edited on VSCode prior to running, and especially saving, on VBA.

<!-- Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script. -->
# Results
## Stock Performance
Only TERP stocks yielded a negative percentage of return in 2017, indicating a bull market. On the other hand, all except ENPH and RUN stocks yielded a negative percentage of return, indicating a bear market. Only RUN stocks increased in percentage of return from 2017 to 2018.

2017 stock performance is shown below.

![2017_ Stock_ Performance](https://user-images.githubusercontent.com/96349090/150025334-b6332c5a-0f30-4f11-a5bd-b4f9f5534868.png)

2018 stock performance is shown below.

![2018_ Stock_ Performance](https://user-images.githubusercontent.com/96349090/150025387-59f14853-4368-4442-9d46-912208c72651.png)

## Refactoring Code and Corresponding Execution time
The refactored code declared additional variables compared to the original. Both codes utilized for loops to calculate and store values of interest, however, the for loop of the refactored code ran through the data once as opposed to the original code where the loop ran through the data several times. In this instance of running the subroutines, AllStocksAnalysis ran slower than AllStocksAnalysisRefactored by approximately 0.52 seconds for analysis of 2017 stocks and by approximately 0.58 seconds for 2018 stocks. Note that subsequent run times may decrease since computer resources are allocated specifically for these tasks.

A portion of code from AllStocksAnalysis is shown below.
![All_ Stocks_ Analysis_ Code](https://user-images.githubusercontent.com/96349090/150013717-31477603-26b7-4496-9c32-d9ef025d0724.png)

A portion of code from AllStocksAnalysisRefactored is shown below.
![VBA_ Challenge_ Code](https://user-images.githubusercontent.com/96349090/150013743-3aa63526-8307-4826-a3fc-d6c2e0105bf1.png)

AllStocksAnalysis ran in roughly 0.67 seconds for year 2017 as shown below.
![All_ Stocks_ Analysis_ 2017](https://user-images.githubusercontent.com/96349090/150008544-5bd5003e-79fb-4053-b878-d9f0bec884fa.png)

ALLStocksAnalysisRefactored ran in roughly 0.15 seconds for year 2017 as shown below.
![VBA_ Challenge_ 2017](https://user-images.githubusercontent.com/96349090/150008664-deec51c7-99bc-41bf-a661-9a894094a8e8.png)

AllStocksAnalysis ran in roughly 0.66 seconds for year 2018 as shown below.
![All_ Stocks_ Analysis_ 2018](https://user-images.githubusercontent.com/96349090/150008593-9a181f7d-ea7f-4047-a5d4-22a2bb603db9.png)

ALLStocksAnalysisRefactored ran in roughly 0.08 seconds for year 2018 as shown below.
![VBA_ Challenge_ 2018](https://user-images.githubusercontent.com/96349090/150008678-8b77a0b9-9ea4-4da8-8022-313c7eb47791.png)

<!-- Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script? -->
# Summary
An advantage of refactoring code is that the refactored code will run in a shorter timespan than the original. The distinction between execution times may be commensurate with the amount of data to be analyzed. A possible disadvantage of refactoring code is that should the lines of code increase, the amount of time spent on refactoring may increase in accordance. If no urgency or time constraint exists, refactoring the original code may be optional rather than a necessity. The efficiency of refactoring code is demonstrated in this project with the execution times of AllStocksAnalysis and AllStocksAnalysisRefactored. Overall this project is a worthwhile exercise of coding in VBA and refactoring code so the possible disadvantage is rendered null.
