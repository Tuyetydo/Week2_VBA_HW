# Week2_VBA_HW

## Overview of Project
Steve is a recent graduate with a finance degree and his parents proudly volunteered to be his first clients. They specifically want to invest in green energy and more specifically, they selected DAQO New Energy Corp (DQ) because of the connection that his parents met at Dairy Queen. But before approving any further actions, Steve pulled together a spreadsheet containing the stock data regarding other green energy stocks to have me analyze. I created the workbook for him with the initial VBA and he loved it but is now looking to expand the dataset to include the entire stock market over the last few years as well as reducing the time to execute it.  


### Purpose
The purpose of this project is to create a refactored VBA workbook for Steve to use and analyze in a reasonable timeframe.  The requested deliverables to be provided to Steve are:

- [x] Deliverable 1: Refactor VBA code and measure performance
- [x] Deliverable 2: A written analysis of my results

## Results and Summary

### Performance Results
The quickest benchmark to compare if the refactoring reduced the execution times, is by comparing the run times between the codes. Below is a comparison of the 2017 pre and post refactoring run times. The pre time shows that the code ran in 1.15 seconds while the post time shows that the code ran in 0.13 seconds. The difference of 1.02 seconds which is almost 89% less run time.

![2017 PreFactoring Image](/Resources/2017%20Run%20Time.PNG)

![2017 ReFactored Image](/Resources/VBA_Challenge_2017.PNG)

Below is a comparison of the 2018 pre and post refactoring run times. The pre time shows that the code ran in 1.17 seconds while the post time shows that the code ran in 0.13 seconds. The difference of 1.04 seconds which is almost 89% less run time. 

![2018 PreFactoring Image](/Resources/2018%20Run%20Time.PNG)

![2018 ReFactored Image](/Resources/VBA_Challenge_2018.PNG)

Beyond looking at the execution times, it's interesting to see the growth and development of the codes and the narrative of the steps of what the VBA are accomplishing. The beginning code focused on locating the starting and ending prices while the refactored code focused on looping but confirming before completing the run.

- Original 
![PreCode](/Resources/PreCode.PNG)

- Refactored
![PostCode](/Resources/PostCode.PNG)

### Summary
The concept of refactoring is reviewing to improve and make it more efficient. Based on the performance results, specifically the run times. It's easily seen the advantage of refactoring the code in this task with a reduced run time of 89%. At the same, with any review there is time dedicated to reviewing and adjusting with no guaranteed that the edits will either improve the run time or even garner the same results in the end. 

### Advantages and Disadvantages 
The advantages of the refactored code is that it allows for not only quicker results, it has check points to ensure that the selected tickers are accounted for. The disadvantages in comparison to the original code is the simplicity of it and the user friendliness of reading the code. 
