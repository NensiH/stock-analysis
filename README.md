# Stock-analysis

## Overview of Project

Steve's parents are passionate about Green Energy and there are many forms of green energy to invest in. They haven't done much research and wanted to invest their money in DAQO new energy corps which makes solar panels, and its Ticker symbol is 'DQ'. So, Steve wants to do much analysis for other stocks options as well as DQ, As DQ's return was less in 2018. Here, we have also refactored code for better and faster results.

## Purpose
he purpose of this project was to make an efficient way to look at multiple stocks using VBA. As Steve's parents were interested to invest in 'DQ', We are helping them in details of stock analysis through VBA code. Here, we are using Stocks data from the year 2017 and 2018. 
Since Daqo might not be the best option for Steve's parents to invest in, we are analyzing multiple stocks to find some better choices for them. Repurposed to analyze any stock. With a little more code, we can analyze a whole list of stocks. 
 
## Analysis and Challenges
 After doing research and Analysis Steve find out that DQ dropped over 63% in 2018 and he want to offer some better stocks to his parents.
 Defining and using multiple variables and for loops we have created a flexible macro for running multiple stocks for 2018 and 2017.
To make this easy to understandable and readable for Steve, we are formatting tables like changing font styles, adding borders, setting number formats, and so on—but we can automate formatting with VBA. Also, Steve wanted a button that would be easier and more user-friendly to run any year Analysis.

Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Here, we are refactoring this code successfully made the VBA script run faster.

### DQ Analysis
Steve wants to know how DQ performed in 2018. One way to measure this is to calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year.
Here is the VBA script I have used for DQ Analysis:

<img width="481" alt="Screen Shot 2021-11-03 at 11 01 49 AM" src="https://user-images.githubusercontent.com/92277581/140097182-3ff56867-7a12-473b-a677-581399938459.png">


### All stocks Analysis
   Since DQ was not the best option Steve wants to check all stocks returns for the years of 2017 and 2018, as well as may want to look at a different set of stocks in the future with the refactoring method and check whether refactoring your code successfully made the VBA script run faster and By doing it this way, the analysis would be completed much faster than using the nested for loop for earlier. 
   
   In order to make my code more efficient, I switch the nesting order of for loops. For this, I created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays with the tickers array by using a variable called the tickerIndex.
   
   Here is the code for comparing the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script:
   
<img width="588" alt="Screen Shot 2021-11-03 at 10 55 42 AM" src="https://user-images.githubusercontent.com/92277581/140096224-127f3a89-1e37-4d4f-a2d3-f4b9ff6f063b.png">


<img width="692" alt="Screen Shot 2021-11-03 at 10 56 25 AM" src="https://user-images.githubusercontent.com/92277581/140096244-d670b025-ebd2-4d7b-97a2-a0d3fbcd8044.png">

<img width="658" alt="Screen Shot 2021-11-03 at 10 57 04 AM" src="https://user-images.githubusercontent.com/92277581/140096259-f97ea555-1c2d-4600-9a3b-b1b49e4947d8.png">




   Here is the Screenshot for VBA script to run faster with refactoring:
   
  <img width="262" alt="Screen Shot 2021-10-31 at 9 37 31 PM" src="https://user-images.githubusercontent.com/92277581/140094155-1bd9da6d-a7d8-44e3-b047-7aefc71f1da2.png">

 <img width="264" alt="Screen Shot 2021-10-31 at 9 37 12 PM" src="https://user-images.githubusercontent.com/92277581/140094169-721eb336-4402-4e6e-b0e1-adba3ecb68a1.png">

   
  
## Challenges and Difficulties Encountered

Tracing and Debugging was little bit hard because I have used many for loops and after making the appropriate adjustments (Refactoring code) to your script that will allow it to run on every worksheet, i.e., every year, just by running it once.

It took a while to fix the problem when changed the code to refactor. Also got many times errors in For loops and defining tikerindex. In conclusion VBA script run faster with Refactoring.

## Results

Here is the results Steve got for 2018 DQ stock Analysis, dropped over 63% in 2018:

<img width="229" alt="Screen Shot 2021-11-03 at 10 27 57 AM" src="https://user-images.githubusercontent.com/92277581/140091154-dfbed542-f464-4267-a325-184ad0cfc857.png">


  Here is the results for 2017 All stocks Analysis, in which only TERP dropped over 7% but all other stocks are up:
   
   <img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92277581/140009891-fb368e0d-67b0-40dd-8c23-8565c6377a1c.png">

Here is the results for 2018 All stocks Analysis and looks like Stock market was down in 2018 only RUN and ENPH :

<img width="247" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92277581/140009944-3f207cf5-38c8-4d8d-80b7-2807a9275951.png">

By looking into both 2017 and 2018 data analysis, RUN and ENPH are better options to invest into as it gives good returns for both years. 
Based on the run-times, the refactored code runs about 0.5 seconds faster than the original code making it more efficient.

## Summary

- What are the advantages or disadvantages of refactoring code?

  **Disadvantages**: As per my analysis, It involves more resources for file/sheets to read and write because it involve both sheets 2017, 2018 and All stocks analysis at the same time. As well as it takes more time to write refactoring code.
 
  **Advantages**: As per my analysis, it uses less file/resources because it involves at a time only 1 sheet for read and write. So, it takes more time to run this code than refactoring one. It is complex while coding, but it improves performance by avoiding multiple resources involved.


- How do these pros and cons apply to refactoring the original VBA script?

1. In the refactoring code, we have taken all the tickers from 2017 and 2018 sheets and then returns to All stock analysis (Simple read and write operation). 

2. In the refactoring code, We have calculated Total volume, ticker starting price and ticker ending prices in memory (With Array variable) and then we did simple write for all this 3 values to the All stock Analysis sheet.

  

