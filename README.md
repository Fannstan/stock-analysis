# Stock-Analysis

## 1. Overview of Project: 
      Steve has been assisting his parents with building a stock portfolio that is comprised of clean energy or green stocks.  Using VBA, the data that Steve provided for 12 green stocks was analyzed and presented in an excel workbook.  The analysis showcased the total daily volume traded and the return for each stock.  The purpose of this project was to take this analysis a step further and refactor the VBA script so that it can be easily used by Steve to analyze any number of stocks for his clients.  The goal of this project is speed up the run time of the VBA script by creating a ticker index and enhancing the code so that the data only needs to be looped through one time to produce the same analysis for each stock.   
      
## 2. Results:
      The original VBA code run time for 2017 was 0.9023438 seconds and for 2018 was 1.003906 seconds, as shown in the screen shots below: 
      
      ![Timer_Original_Script_2017](https://user-images.githubusercontent.com/103215123/166998340-8d9a854f-0604-4f18-8efb-0414b77f8551.png)

      ![Timer_Original_Script_2018](https://user-images.githubusercontent.com/103215123/166998619-26815504-114e-4c53-abcc-a1a35995d18e.png)

      
      In the original VBA code, we created an outer for Loop (i) that set the total volume to zero after each time the code looped or finished through the analysis of a  ticker, the inner for Loop (j) runs through each row giving the total volume and the starting and ending prices of each ticker to calculate the return.  
      
      ![Original_Nested_Loop](https://user-images.githubusercontent.com/103215123/166984053-ac76ae3d-802b-4853-973f-40df1ba54a32.png)
      
      The Refactored VBA code uses a ticker index, and sets the total volume, starting price, and ending price as output arrays.  THe refactored code also initializes the ticker volumes to zero.  The second loop goes over all of the rows in the spreadsheet, but uses the ticker index to only loop through the rows one time to get the outputs for ticker, volume, and return.  
      
      ![Refactored_tickerIndex_Arrays](https://user-images.githubusercontent.com/103215123/166985717-fc91e494-1d70-4cd7-83b4-cdf736edac2e.png)

      Below is the code showing how the ticker index was used to help the code run more efficiently: 
      
      ![Refactored_tickerIndex_Loop](https://user-images.githubusercontent.com/103215123/166986715-0a357bee-e19c-4b43-93a5-c5b3bf83df80.png)

      The result of refactoring the code is significantly shorter run times for each year analyzed.  The refactored code for 2017 ran in 0.1367188 seconds, almost 0.77 seconds faster.  The refactored code for 2018 ran in 0.1367188, almost 0.87 seconds faster.  
      
## 3. Summary: 
      3a) What are the advantages or disadvantages of refactoring code?
      Refactoring code can improve the efficiency of the code and make the code easier to follow for other programers to use in the future. Some disadvantages of refactoring are that though the code may run faster and be easier to read or understand for programmers who use the code in the future, these advantages are not always seen by the end user.  
      
      3b) How do these pros and cons apply to refactoring the original VBA script?
      The pros of refactoring the code in this example is that the refactored code has faster run times than the original code and it has been reorganized in a way that makes it more applicable to larger amounts of data. The biggest disadvantage for refactoring this project was the amount of time taken to rewrite the code.  
