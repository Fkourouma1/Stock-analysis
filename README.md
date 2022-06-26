# Stock-analysis
# Overview of Project
The main purpose of this project is to help Steve's parents to analyse All 12 stocks in our dataset for previous years.We are also helping Steve to run the code for faster execution by refactoring our code in the VBA. We will end the project by writting an analysis that explains in details everything we did. 
## Stock Performance
### 2017
In 2017, we have noticed that all stocks performed well except TERP which return was down -7.2%. The most traded stock was SPWR with 23.1% as performance. The stock with the highest performance was DQ as the return was 199.4%. We calculated the return by using this script : "Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1". The original script run in ............. while the refactured script run in 0.140625 seconds. We did the formating by assigning the positive return to a green color and negative to a red color. The formating was done by using the script " If Cells(i, 3) > 0 Then Cells(i, 3).Interior.Color = vbGreen Else Cells(i, 3).Interior.Color = vbRed" . 
### 2018
The following year 2018, All stocks underperformed except ENPH with a return of 81.9% and Run with the highest return as 84%. The original script run in ............. while the refactured script run in 0.1523438 seconds. Reason why ....................... We formated the header in bold by using " Range("A3:C3").Font.FontStyle = "Bold" " formala
## Summary Statement 
### Advantages and disadvantages of refactoring a code
The main advantage of refactoring a code is to facilitate the comprehension and extendibility of our project Which could help the code run faster. Refactoring help understand and read the code easily 
### Disadvantages of refactoring a code 
On the other side the disadvantage of refactoring is that if the code is imprecise, it could bring some confusion which will bring new bugs and erros into the code. 
### Pro and cons apply of refactoring compare to the orignal script
......................
