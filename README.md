# Stock-analysis
# Overview of Project
The main purpose of this project is to help Steve's parents to analyse All 12 stocks in our dataset for previous years.We are also helping Steve to run the code for faster execution by refactoring our code in the VBA. We will end the project by writting an analysis that explains in details everything we did. 
## Stock Performance
### 2017
In 2017, we have noticed that all stocks performed well except TERP which return was down -7.2%. The most traded stock was SPWR with 23.1% as performance. The stock with the highest performance was DQ as the return was 199.4%. We calculated the return by using this script : "Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1". The original script run in 1.273438 seconds while the refactured script run in 0.140625 seconds. We did the formating by assigning the positive return to a green color and negative to a red color. The formating was done by using the script " If Cells(i, 3) > 0 Then Cells(i, 3).Interior.Color = vbGreen Else Cells(i, 3).Interior.Color = vbRed" . 
### 2018
The following year 2018, All stocks underperformed except ENPH with a return of 81.9% and Run with the highest return as 84%. The original script run in 1.273438 seconds while the refactured script run in 0.1523438 seconds. We formated the header in bold by using " Range("A3:C3").Font.FontStyle = "Bold" " formula
## Summary Statement 
### Advantages and disadvantages of refactoring a code
The main advantage of refactoring a code is to facilitate the comprehension and extendibility of our project Which could help the code run faster. Refactoring help understand and read the code easily 
### Disadvantages of refactoring a code 
On the other side the disadvantage of refactoring is that if the code is imprecise, it could bring some confusion which will bring new bugs and erros into the code. 
### Pros and cons apply of refactoring compare to the orignal script
#### Pros of refactoring 
In the refactored code we intialized the tickerindex to 0 to make it easier to read while in the original script we just set the variable starting and ending prices. When incresing the volume for current ticker, in the refactored script  we made it clear and consise by assigning tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value. while in the original script we did not condense as the formula shows : If Cells(j, 1).Value = ticker Then totalVolume = totalVolume + Cells(j, 8).Value. Another proof of refactoring being more understandable is that when getting the ticker ending and starting prices in the refactored we use this formula If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(j, 6).Value but in the original script we used If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then startingPrice = Cells(j, 6).Value, which seems to be a little complex. 
#### Cons of refactoring 
In one of our example, when creating a loop to initialize the tickerVolumes to zero, we use this formula in the refactored script   For i = 0 To 11 tickerVolumes(i) = 0 which can sometimes bring some confusion since it is imprecise. while the original script shows a longueur script but well detailed as the formula was For i = 0 To 11 ticker = tickers(i) , totalVolume = 0
