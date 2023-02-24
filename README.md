# **Automated Stock analysis with VBA**
## **Teodor Anderson**

## Overview of Project
	
Steve, a good friend of mine , graduated with a finance degree and acquired his first clients, his ecologically minded parents, and works to build a diversified portfolio of green energy stocks. He has enlisted my help to analyze a data set of ecological stocks from 2017 and 2018 to find the most successful options. Using VBA to automate excel analysis, my task is to write and refactor macros to create dynamic input buttons to extract data from Excel and  output the total daily volume and its yearly return into clean, visually appealing data back to Excel with the lowest runtime as possible.
  

# Results

# Unfactored Code
  ![Undone Macro Top](https://user-images.githubusercontent.com/116928193/204114534-3fc38e8a-1611-40e5-80cf-3c5851936d91.png)
  ![Unfactored middle](https://user-images.githubusercontent.com/116928193/204114538-cec944c7-a5c3-4d58-b4de-c2662dd512e1.png)
  ![Undone Macro Bottom](https://user-images.githubusercontent.com/116928193/204114545-cae4014f-ef0f-4d5b-b3ab-4387fd262125.png)
   ![2018 bad](https://user-images.githubusercontent.com/116928193/204114640-d485874b-9a72-407e-93e9-ac0792a788cb.png)
    ![Bad Runtime 2017](https://user-images.githubusercontent.com/116928193/204114647-6ba95194-499f-4c9a-933a-9450a6c839ca.png) 

## Analysis of Unfactored code

As seen in the code above, the language used in the unfactored code uses many "code smelly" functions, kludges like unnamed "i" and "j" variables that could make the next person or even the author extremely confused at the functions and purpose of the code. The unfactored code uses a single array to store tickers and runs through the loop outputting each value as the loop progresses, which in coding is very slow, and as you will see later, the factored code stores all the values into arrays made for every out put and then after the for loop is done it will use the output arrays to output the data.
	
	
  # Factored code
![Factored Code 1](https://user-images.githubusercontent.com/116928193/204118413-ef371572-6af0-42c5-9b0c-bb3af88fdfa1.png)
![Factored Code 2](https://user-images.githubusercontent.com/116928193/204118423-cea91d19-ed33-4ceb-8596-f41d11d6c1c3.png)
![VBA_Challenge_2018 png - Copy](https://user-images.githubusercontent.com/116928193/204118429-ee5b604c-a36f-4040-8076-15c2edf7be88.png)
![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/116928193/204118432-d9ea24aa-f6a8-4d61-84c8-15c907ea353d.png)


  ## Analysis of Factored Code
By creating output arrays for volume, starting and ending prices, as well as creating for loops to increase the volume by each array ticker and increase the ticker for every loop. the code runs much faster on average more than 4 times as fast, also the code is much cleaner, frequently annotated for future changing or inspecting and free of kludges having every variable labeled for its function. The code doesnt have to use one for loops to both loop through the data, check for the correct ticker and output the data, making the code much more streamlined and understandable.
 
# Summary

## Refactoring Code in General

When refactoring any code, the new code will use less memory, speeding up the code at hand as well as any other functions that may need to run along side your code. Another beneficial effect of factoring code is the language becoming much more understandable, making future edits or inspections much easier. As a coder, a large portion of the job is refactoring code, defining values for their functions and writing argument and variable based functions to reduce static and clunky code. 
The disadvantages of refactoring code would be man hours and capital lost in applying skilled coders on the task, or the code being down or buggy while it is being refactored.


## Refactored Code in Vba
  
The refactored code at hand's advantages include not using as much RAM memory as well as being much, up to 5 times faster to run.
The disadvante would be if the dataset would involve more arrays, storing the tickers in three arrays could run into overflow or memory problems, or introducing errors or bugs while refactoring the code.
