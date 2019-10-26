Sub VBAchallenge():

Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate

    Dim LastTicker As String
    LastTicker = ws.Cells(2, 1)
    Dim CurrTicker As String
    Dim TickOpen As Double
    TickOpen = ws.Cells(2, 3)
    Dim TickClose As Double
    Dim TickChge As Long
    Dim LastRow As Long
    Dim LastCol As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    Dim iRowSum As Long
    Dim iColSum As Long
  

    iRowSum = 2
    iColSum = LastColumn + 2
    SumTickCol = LastColumn + 2
    SumYrChgeCol = LastColumn + 3
    SumPerChgeCol = LastColumn + 4
    SumTotVolCol = LastColumn + 5
    GreatestTypCol = LastColumn + 8
    GreatIncTickCol = LastColumn + 9
    GreatValCol = LastColumn + 10


    ws.Cells(1, SumTickCol) = "Ticker"
    ws.Cells(1, SumYrChgeCol) = "Yearly Change"
    ws.Cells(1, SumPerChgeCol) = "Percent Change"
    ws.Cells(1, SumTotVolCol) = "Total Stock Volume"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
  
        For i = 2 To LastRow + 1

                CurrTicker = ws.Cells(i, 1)
                CurrVol = ws.Cells(i, 7)
                
                If CurrTicker = LastTicker Then
                TickVol = TickVol + CurrVol
                TickClose = ws.Cells(i, 6).Value
                Else

                ws.Cells(iRowSum, iColSum) = LastTicker
                ws.Cells(iRowSum, iColSum + 1) = TickClose - TickOpen
                
                If ws.Cells(iRowSum, iColSum + 1) >= 0 Then
                ws.Cells(iRowSum, iColSum + 1).Interior.ColorIndex = 4
                Else: ws.Cells(iRowSum, iColSum + 1).Interior.ColorIndex = 3
                End If

                If TickOpen = 0 Then
                ws.Cells(iRowSum, iColSum + 2) = 0
                Else
                ws.Cells(iRowSum, iColSum + 2) = (TickClose - TickOpen) / TickOpen
                End If
                
                ws.Cells(iRowSum, iColSum + 3) = TickVol
                TickOpen = ws.Cells(i, 3).Value
                TickClose = ws.Cells(i, 6).Value
                TickVol = ws.Cells(i, 7).Value
                iRowSum = iRowSum + 1
                LastTicker = ws.Cells(i, 1)
                End If
 

        Next i
        
                lSumRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                ws.Range("I2:L" & lSumRow).Sort key1:=ws.Range("L2:L" & lSumRow), order1:=xlDescending, Header:=xlNo
                ws.Range("P4") = ws.Range("I2")
                ws.Range("Q4") = ws.Range("L2")
                
                ws.Range("I2:L" & lSumRow).Sort key1:=ws.Range("K2:K" & lSumRow), order1:=xlDescending, Header:=xlNo
                ws.Range("P2") = ws.Range("I2")
                ws.Range("Q2") = ws.Range("K2")
                
                ws.Range("I2:L" & lSumRow).Sort key1:=ws.Range("K2:K" & lSumRow), order1:=xlAscending, Header:=xlNo
                ws.Range("P3") = ws.Range("I2")
                ws.Range("Q3") = ws.Range("K3")
                
                ws.Range("I2:L" & lSumRow).Sort key1:=ws.Range("I2:I" & lSumRow), order1:=xlAscending, Header:=xlNo
                                
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(3, 17).NumberFormat = "0.00%"
                ws.Columns("A:Z").AutoFit
                ws.Range("K:K").NumberFormat = "0.00%"
Next ws

End Sub
