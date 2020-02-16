'Create a script that will loop through all the stocks for one year for each run and take the following information.

'The ticker symbol.

'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock.


Sub Totalstockvolume():

'Loop through these worksheets
For Each ws In Worksheets

'Define my variables

Dim i, j As Long
Dim totalvolume As Double
Dim ticker As String
Dim lastrow As Long
Dim yearlychange As Double
Dim percentage As Double
Dim openprice As Double
Dim currentprice As Double
Dim openprice_row As Long

'Add Header Name to Display Data
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Value"

 'Set Initial Total
 totalvolume = 0
 j = 2
 openprice_row = 2

 'Determine the last Row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'Each Year of Stock Data
 For i = 2 To lastrow
     
     'Compare Each Ticker
     If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then

         'Calculate Total Volume
         totalvolume = totalvolume + ws.Range("G" & i).Value

     Else
         'Ticker price change
         ticker = ws.Range("A" & i).Value

         'Calculate Yearly Change and Percent Change
         openprice = ws.Range("C" & openprice_row)
         currentprice = ws.Range("F" & i)
         yearlychange = currentprice - openprice

         'Calculate Percent Change
         If openprice = 0 Then
            percentage = 0
         Else
            percentage = yearlychange / openprice
         
    End If

         'Range Title Ticker,Total Volume,Yearly Change and Percent Change
         ws.Range("I" & j).Value = ticker
         ws.Range("L" & j).Value = totalvolume + ws.Range("G" & i).Value
         ws.Range("J" & j).Value = yearlychange
         ws.Range("K" & j).Value = percentage
         ws.Range("K" & j).NumberFormat = "0.00%"
         
         'Conditional Formating Yearly Change, Positive Green / Negative Red
         If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
         Else
            ws.Range("J" & j).Interior.ColorIndex = 3
         
    End If

        'Cell Labels for Challenge Part
        ws.Cells(2, 15).Value = "Greatest Total Volume"
        ws.Cells(3, 15).Value = "Greatest Percent Increase"
        ws.Cells(4, 15).Value = "Greatest Percent Decrease"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Total Value"
        
    'Display to new sheet
         
    j = j + 1
        
    totalv = 0
        
    openprice_row = i + 1
         
     End If
 
 Next i
 
 Next ws

End Sub