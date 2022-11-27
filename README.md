# **Module 2 Assignment: Green Stock Analysis**
## **Teodor Anderson**

## Overview of Project
	
Steve, a good friend of mine , graduated with a finance degree and acquired his first clients, his ecologically minded parents, and works to build a diversified portfolio of green energy stocks. He has enlisted my help to analyze a data set of ecological stocks from 2017 and 2018 to find the most successful options. Using VBA to automate excel analysis, my task is to write and refactor macros to extract data from Excel, create dynamic input interactions, and output the total daily volume and its yearly return into clean, visually appealing data back to Excel as quickly as possible.
  

# Results

# Unfactored Code
  ![Undone Macro Top](https://user-images.githubusercontent.com/116928193/204114534-3fc38e8a-1611-40e5-80cf-3c5851936d91.png)
  ![Unfactored middle](https://user-images.githubusercontent.com/116928193/204114538-cec944c7-a5c3-4d58-b4de-c2662dd512e1.png)
  ![Undone Macro Bottom](https://user-images.githubusercontent.com/116928193/204114545-cae4014f-ef0f-4d5b-b3ab-4387fd262125.png)
   ![2018 bad](https://user-images.githubusercontent.com/116928193/204114640-d485874b-9a72-407e-93e9-ac0792a788cb.png)
    ![Bad Runtime 2017](https://user-images.githubusercontent.com/116928193/204114647-6ba95194-499f-4c9a-933a-9450a6c839ca.png) 

## Analysis of Unfactored code

As seen in the code above, the language used in the unfactored code uses many "code smelly" functions, kludges like unamed "i" and "j" variables that could make the next person or even the author extremely confused at the functions and purpose of the code. The unfactored code uses a single array to store tickers and goes through the loop outputting each value as the loop progresses, which in coding is very slow, and as you will see later, the factored code stores all the values into arrays made for every out put and then after the for loop is done will output the one by one.
	
	
  # Factored code
![Factored Code 1](https://user-images.githubusercontent.com/116928193/204118413-ef371572-6af0-42c5-9b0c-bb3af88fdfa1.png)
![Factored Code 2](https://user-images.githubusercontent.com/116928193/204118423-cea91d19-ed33-4ceb-8596-f41d11d6c1c3.png)
![VBA_Challenge_2018 png - Copy](https://user-images.githubusercontent.com/116928193/204118429-ee5b604c-a36f-4040-8076-15c2edf7be88.png)
![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/116928193/204118432-d9ea24aa-f6a8-4d61-84c8-15c907ea353d.png)


  ## Analysis of Factored Code
  
