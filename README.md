# stock-analysis
Analysis of Steve s stocks
Week 2 of the Data Analysis Boot Camp

# 1. Introduction

This challenge is done as part of the second week UCSD boot camp to build on the skills we have learned in this module especially coding skills in VBA.
The code is refactored and perfromed analysis as per the requirements and also considered the advantages and disadvantages of refactoring. 

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

This project uses VBA to take stock ticker data from an existing spreadsheet, analyze the data, and summarizes the spreadsheet's results. The final step is saving the results to an output directory so that the original data remains untouched.

## 2. Files

In this assignment it consists of one technical deliverable, resource folder with input, performance pop up images and a written report to deliver your results. 
  
## 3. Analysis
### Overview of Project

The purpose of this project is we are using VBA code to help us analyze the Stocks for Steve's parents so that they could invest in right business.
Our goal is loop through each ticker and find their total volume over the year interested.Currently the available data has two worksheets for the years 2017 and 2018.

Using the VBA code we looped through each ticker and calculated the total volume, starting price, final price of each stock and placed it in arrays. 
These array values are shown in a seperate worksheet in excel with proper headers.

Also in the code the "Returns" column is formatted to show Steve which stocks have (increased -Green color or decreased - Red Color)shown positive results based on the final and starting prices. To do this the code has been refactored to add them all in arrays so that the resultant table has every ticker details.The code has been refactored to fit in these details.

While refactoring, the code has been tested for performance using the Timer functions in VBA code. At the end we made sure the code is refactored to the requirements by avoiding addition of new code and made it efficient so that in future this code works better if new tickers are added.

The code has interactive buttons to run the analysis considering the year as input.

![VBA_Challege_InputBox](https://user-images.githubusercontent.com/111100908/186755697-ea139298-ee4b-4717-a36b-bbcf91b8055d.png)


### Results: 

## After comparing the results of 2017 and 2018 stocks its clearly understood that year 2017 most of the stocks have out performed comparing to the year 2018. The tables in the images clearly shows that. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/111100908/186755700-9a0dcba7-f849-44fc-b8c8-9a8fb8a5228c.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111100908/186755702-64d31833-e1c3-438f-bcfc-e910c06cfb00.png)


## Refactored Script :The code that has been refactored to accomodate all the tickers is shown below:
![VBA_Challenge_Refactor code](https://user-images.githubusercontent.com/111100908/186974193-179048c6-d50d-454d-b819-87a44c52424c.png)


### Summary: 

  #### What are the advantages or disadvantages of refactoring code?
  
     - Refactoring improves the design of software, increases the performance measure and gives a scope for the code extensibility.
     - Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; 
       you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic
       of the code to make it easier for future users to read the code. 
     - In this challenge, VBA code module is very small  so the refactoring code may or may not show a big difference 
       but when the application is too big its not worth refactoring. It might costs the time and loses track of requirements 
       and most importantly readability decreases to the future developers.
    
  #### How do these pros and cons apply to refactoring the original VBA script?
  
     - The design has been included with all the tickers in the original worksheet. 
     - If the user input is valid then only the analysis is started which includes the required initialisations (memory and compilation time considerations).
     - The readability of coding might gets lost if the conditions are not explained well using comments.
        
              
