<h1> <p align=center>Stocks Analysis Report</p> </h1>


<h2> <p align=center>Project Overview</p> </h2>

<h3><p align=left>Background </p></h3> 
A stock analysis application was built using Excel-VBA to help Steve analyze stock dataset. The stock dataset is mainly comprised of a handful of renewable energy stocks. Along with expanding current functionality of Steve's current stock analysis workbook, the VBA code in the workbook was refactored. Refactored code enabled Steve to analyze the stock data for multiple years with some added benefits, discussed later in detail.

<h3><p align=left>Purpose</p></h3>
This report provides a detailed overview of the refactoring process and analyzes the stock performance of years 2017 and 2018.

<h2> <p align=center>Results</p> </h2>

<h3><p align=center>Stock Performance Analysis</p></h3>

All Stocks Analysis 2017   |  All Stocks Analysis 2018
:-------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/90424752/139604743-e4e549fa-41fc-4077-ba62-47c03a65243b.png)  |  ![](https://user-images.githubusercontent.com/90424752/139604748-68483e12-afe7-4bd0-8bb7-ba35363bb632.png)

The dataset is mainly comprised of 12 select green energy stocks.
* From the analysis results, 11 out of 12 stocks gave positive returns in the year 2017, indicating overall green energy market was up.
* DQ performed extraordinarily well in year 2017 giving 199.45% return on investment, followed by SEDG with 184.47% return on investment.

However, the scenario for green energy industry changed drastically in 2018.
* In 2018 only 2 out of 12 green energy stocks were in profit, namely RUN (83.95% return) and ENPH (81.92% return).
* DQ was on -62.6% in loss as compared to its value at the start of year 2018.

**Conclusions & Recommendations:**
* As per the above analysis, if Steve's clients were to invest only in the green energy industry, RUN and ENPH seem to be better investment options, as both stocks gave positive returns over the entire period of analysis.

<h2><p align=center>Run Time Analysis </p></h2>

Let us take a look at the structural changes made to the initial code in order to refactor it.

**Features of the original (un-refactored) code:**
* The initial code used Nested for loops, one for iterating through the Tickers array (containing 12 ticker names) and another iterating through all the rows of the dataset.
* The initial code used four variables ticker, totalVolume, startPrice and endPrice to store all the required values. These variables were then used to output the stored information at the end of each iteration, so that they could store new values during the next iteration.
* The main structure of the initial code can be referred to in the code block below.

<pre>
 '4) Outer For loop to loop through the tickers
  <b>For i = 0 To 11 </b>
        
            ticker = tickers(i)
            totalVolume = 0
        
            '5) loop through rows in data
            Worksheets(yearValue).Activate
        <b> For j = 2 To RowCount </b>
                    
                        '5) a) find total vol for the current ticker
                        If Cells(j, 1).Value = ticker Then
                            totalVolume = totalVolume + Cells(j, 8).Value
                        End If
                    
                        '5) b) find starting price for the current ticker
                        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                            startPrice = Cells(j, 6).Value
                        End If
                
                        '5) c) find ending price for the current ticker
                        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                            endPrice = Cells(j, 6).Value
                        End If
                        
    <b>     Next j </b>
    
            '6) output the data for the current ticker
            Worksheets("All_Stocks_Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endPrice / startPrice - 1
                
<b> Next i </b>
</pre>
**Features of the refactored code:**
* The refactored code was restructured to remove the Nested for loop. Instead, there was only one loop iterating through all the rows.
* To replace the for loop iterating through the tickers, a new variable tickerIndex was introduced. The tickerIndex variable was used to refer to the different tickers in the datatset. 
* This restructuring also required three additional arrays to store the values of ticker volumes, starting and ending prices of the stocks. these three arrays were later used to retrieve and display all the values in the analysis table.

<pre>

 <b> Three additional arrays to store values
    Dim tickerVolumes(0 To 11) As Long
    Dim tickerStartingPrices(0 To 11) As Single
    Dim tickerEndingPrices(0 To 11) As Single </b>
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
  <b>''2b) Loop over all the rows in the spreadsheet.
     For i = 2 To RowCount </b>
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
  <b>Next i </b>

</pre>

We can see in the screenshots below that the code **run time of the refactored code for both the analyses has gone down significantly**. This is **mainly due to the number of reduced operations achieved by getting rid of the nested for loops**. 
This restructuring of the code was mainly possible due to certain characteristics of the dataset. These characteristics have been discussed in detail in the further sections.


Run-time Before Refactoring |  Run-time for Refactored Code
:--------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/90424752/139733149-5a3d7b64-9b54-4477-a431-c850c00a5370.png)|![](https://user-images.githubusercontent.com/90424752/139733213-51941a56-94cf-489d-9751-d9d872a1fc0e.png)

Run-time Before Refactoring|  Run-time for Refactored Code
:-------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/90424752/139734117-798c222f-0c99-40eb-80c5-168151ae40c2.png)|![](https://user-images.githubusercontent.com/90424752/139736243-5180cd2e-3141-4603-aadb-9567f5417ba2.png)


<h2> <p align=center>Project Summary</p> </h2>

<h3> <p align=left>Code Refactoring In General</p> </h3>
Code refactoring is a process of restructuring an existing body of code, altering its internal structure without changing its external behaviour.


<h3> <p align=left>Code Refactoring In General: Pros </p> </h3>

Some of the advantages of refactoring code are as follows:
* Code refactoring is intended to improve the design and structure of the code while preserving its functionality.
* Refactoring of code may improve code readability and reduce complexity.
* Refactoring is usually targeted towards increasing efficiency by decreasing the time taken for code execution.
* Depending on the objective, the refactored code could use less memory hence making the code less resource dependent.
 

<h3> <p align=left>Code Refactoring In General: Cons </p> </h3>

Some of the disadvantages of refactoring code can be:
* Refactoring of the code is generally a time consuming process.
* Sometimes, refactoring code for a certain task could make it more specific for that task and can restrict it's use to particular tasks or situations.



<h3> <p align=center>Code Refactoring : As It Applies To The Current Analysis </p> </h3>

<h4>A Small Recap of Restructuring Mechanism :</h4>
Refactoring of the code in this analysis mainly involved getting rid of a Nested for loop and many extra operations the Nested for loop was performing.
The code was restructured by introducing three additonal arrays and a tickerIndex variable. This tickerIndex variable was used to reference to the tickers in the dataset.

However, this particular restructuring of code was possible because of certain characteristics of the dataset. 
* The dataset was previously sorted in the ascending order of the ticker names.
* The Tickers array that contained ticker names, folllowed the exact same order of the names as in the dataset. 
* Because the order of the tickers was same in the dataset and the ticker array, the condition in the If contidional where we check if the current ticker name is same as that of the dataset, became unnecessary. Removing this condition saved us some more time.

<pre>
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1) <> tickers(tickerIndex) And <b>Cells(i, 1) = tickers(tickerIndex)</b> Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If <b>Cells(i, 1) = tickers(tickerIndex)</b> And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
  </pre>


<h3>Pros & Cons of Code Refactoring As It Applies To The Current Analysis:</h3>

  - **Pros:** Refactored code runs much faster, performing fewer calculations.

  - **Cons:** Refactored code requires 3 additional arrays to store the results during its execution, consuming a lot more resources as compared to the original code.
 

<h3>Pros And Cons Of The Original Code:</h3>

  - **Pros:** Original code uses less memory(resources) while executing, as it uses only 3 variables to store and output the results.

  - **Cons:** It performs many unnecessary calculations while running nested for loop and takes more time for execution.

