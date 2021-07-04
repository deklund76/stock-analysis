# Stock Analysis

## Overview

This repository is an analysis of a handful of green energy stocks. It contains an Excel spreadsheet with data on said stocks, and macros built to analyze and format the data. Some clients looking to invest in green energy decided they wanted to put all their funds into DAQO New Energy Corporation (Stock ticker "DQ"). The following analysis tries to determine whether that is a sound idea and also assesses the run time of the VBA scripts used to analyze the data and the pros and cons of refactoring those scripts from previous scripts (also contained in the macro-enabled workbook).

## Results

Below are the results of running the macros (top) and refactored macros (below).

![2017 results & runtime](https://github.com/deklund76/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018 results & runtime](https://github.com/deklund76/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

![2017 results & runtime (refactored)](https://github.com/deklund76/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Refactored.png)

![2018 results & runtime (refactored)](https://github.com/deklund76/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactored.png)

At a glance, the first thing that's obvious from the data is that 2017 was a very good year for green energy stocks, and the same cannot be said for 2018. 2018 saw all of the selected stocks yield negative returns except for ENPH and RUN which had similar performances. When we look at the big picture, ENPH is clearly the highest performing stock between 2017 and 2018 as it's the only stock to maintain double digit returns for both years. 
Does that mean it will keep going up? Not necessarily. We only have the data in front of us to work with, but a more detailed and nuanced analysis of stock trends and where they're headed is necessary to provide effecive investment advice. The growth in ENPH's stock may not be sustainable for various reasons, in a worst-case scenario we may be investing at it's peak value only to see its value plummet. As for the VBA scripts, both did their jobs, but the refactored script managed to do it in roughly 1/3rd as much time. This is even more impressive when you consider the refactored script included formatting!

## Summary

Why did the refactored script perform so much better? Let's look at the code.

![Original VBA Script](https://github.com/deklund76/stock-analysis/blob/main/Resources/VBA_Script.png)

![Refactored VBA Script](https://github.com/deklund76/stock-analysis/blob/main/Resources/VBA_Script_Refactored.png)

The original script had both nested for loops and conditionals. It was going through every row in the data sheet once _for each different stock_ we only had 12 stocks to look through. In a real world scenario with hundreds or even thousands of stocks trading on the market, this small difference would have ground our program to a screeching halt. By using an if statement to increment our ticker index, we were able to get rid of the outer for loop and improve the program's efficiency _exponentially._ 
So what are the pros and cons of refactoring? Clearly, in this case, refactoring made the code far more efficient, improving the run-time by a factor of 3 while also adding scripts that had been contained in seperate modules making for fewer modules to wade through. On the other hand, we spent valuable coding time to save roughly _2 tenths of a second_ on each use. 
This illustrates the primary pros and cons of refactoring code in general. When working at scale, more efficient code can yield huge benefits in runtime that can be critical in many situations and may even be necessary for software to run without timing out. The trade-off then is the time spent working on refactoring the code and programmers' time isn't cheap! In this case, the original code was very inefficient, but when code has been developed and used over a long time there are often diminishing returns on refactoring. All of these things, the cost to refactor, the benefits of increased runtime, and the state of the original code, need to be weighed when deciding if code should be refactored.

