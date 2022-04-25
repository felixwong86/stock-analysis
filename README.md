# Refactor VBA Code for Stock Analysis

## Purpose
The purpose of this challenge was to refactor (or edit) an existing VBA code from the Module 2 solution and to loop through all of the data one time in order to collect the same data more efficiently (i.e having faster run time speeds). This is achieved through taking fewer steps, using less memory, and/or improving the logic of the code.

## Results
When refactoring the VBA_Challenge.vbs code, there were comments to create a tickerIndex variable and three additional arrays (ticker array already exists). tickerIndex is to be used to access the correct index across the four different arrays:  ticker, tickerVolumes, tickerStartingPrices, and tickerEndingPrices

![Step_1](https://github.com/felixwong86/stock-analysis/blob/main/Resources/VBA_step1.PNG)

A for loop was then created to initaize the tickerVolumes to zero. A script was nested into the loop to increase the current tickerVolumes (stock ticker volume) variable and add the ticker volume for the current stock ticker. Conditional statements were added to check if the current row is the first and last rows with the selected tickerIndex to display the respected variables.
![Step_2](https://github.com/felixwong86/stock-analysis/blob/main/Resources/VBA_step2.PNG)


The result of refactoring the code was an improvement of run time speeds for both 2017 and 2018 analysis outputs compared to the original VBA code.

### Before
![greenstocks_2017](https://github.com/felixwong86/stock-analysis/blob/main/Resources/greenstocks_2017.PNG)
![greenstocks_2018](https://github.com/felixwong86/stock-analysis/blob/main/Resources/greenstocks_2018.PNG)

### After
![VBA_Challenge_2017](https://github.com/felixwong86/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2017](https://github.com/felixwong86/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
