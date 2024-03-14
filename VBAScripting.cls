VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockAnalysis()

    For Each Current In Worksheets
        With Current
            ' Set an initial variable for holding the ticker name
            Dim Ticker As String
            Dim totalTicker As Double
            totalVolume = 0
            Dim endRow As Integer
            endRow = 2
            
            ' Seting headers for the outputs
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            
            ' number of rows in sheet
            Dim Row_Count As Long
            RowRange = .Range("A1").End(xlDown).Row
            
            ' Initialize variables to track maximum values
            Dim maxPercentIncrease As Double
            Dim maxPercentDecrease As Double
            Dim maxTotalVolume As Double
            Dim tickerMaxPercentIncrease As String
            Dim tickerMaxPercentDecrease As String
            Dim tickerMaxTotalVolume As String
            
            ' Loop through all ticker data
            For i = 2 To RowRange
            
                ' Check the ticker
                Dim val1 As String
                Dim val2 As String
                Dim stockOpen As Double
                Dim stockClose As Double
                val1 = .Cells(i + 1, 1).Value
                val2 = .Cells(i, 1).Value
                
                ' first row
                If i = 2 Then
                    ' set first open the open ticker
                    stockOpen = .Cells(i, 3).Value
                End If
                
                If val1 <> val2 Then
                    
                    ' Set the Brand name
                    Ticker = .Cells(i, 1).Value
                    
                    ' Add to the Brand Total
                    totalVolume = totalVolume + .Cells(i, 7).Value
                    
                    ' Last close to calc yearly and perc change
                    stockClose = .Cells(i, 6).Value
                    
                    ' Print the Ticker name in the Summary Table
                    .Range("I" & endRow).Value = Ticker
                    
                    ' Calculate yrly change
                    .Range("J" & endRow).Value = stockClose - stockOpen
                    
                    ' Calculate percentage change only if starting value was not 0
                    ' 0 starting value would simultaneously cause divide by zero error
                    ' and be meaningless
                    If (stockOpen <> 0) Then
                        .Range("K" & endRow).Value = ((stockClose - stockOpen) / stockOpen) * 100
                    End If
                    
                    ' Print the Ticker Total Volume to the Summary Table
                    .Range("L" & endRow).Value = totalVolume
                    
                    ' Increment the row/add
                    endRow = endRow + 1
                    totalVolume = 0
                    stockOpen = .Cells(i + 1, 3)
                    
                    ' Calculating greatest % increase and decrease
                    If .Range("K" & endRow).Value > maxPercentIncrease Then
                        maxPercentIncrease = .Range("K" & endRow).Value
                        tickerMaxPercentIncrease = Ticker
                    End If
                    
                    If .Range("K" & endRow).Value < maxPercentDecrease Then
                        maxPercentDecrease = .Range("K" & endRow).Value
                        tickerMaxPercentDecrease = Ticker
                    End If
                    
                    If .Range("L" & endRow).Value > maxTotalVolume Then
                        maxTotalVolume = .Range("L" & endRow).Value
                        tickerMaxTotalVolume = Ticker
                    End If
                    
                ' If the cell immediately following a row is the same ticker...
                Else
                
                    ' Add to the Ticker Total Volume
                    totalVolume = totalVolume + .Cells(i, 7).Value
                    
                End If
                
            Next i
            
            ' Apply conditional formatting directly to the range
            .Range("J2:J" & endRow - 1).FormatConditions.Delete
            .Range("J2:J" & endRow - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
            .Range("J2:J" & endRow - 1).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
            .Range("J2:J" & endRow - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
                Formula1:="=0"
            .Range("J2:J" & endRow - 1).FormatConditions(2).Interior.Color = RGB(0, 255, 0)
            
            ' Output the stocks with the greatest percent increase, decrease, and total volume
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            .Range("P2").Value = tickerMaxPercentIncrease
            .Range("P3").Value = tickerMaxPercentDecrease
            .Range("P4").Value = tickerMaxTotalVolume
            .Range("Q2").Value = maxPercentIncrease
            .Range("Q3").Value = maxPercentDecrease
            .Range("Q4").Value = maxTotalVolume
            
        End With
    Next

End Sub

