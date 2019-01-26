Attribute VB_Name = "Module1"
Sub TickerTally()

Rem Declare Variables

Dim WorksheetName As String

Dim NewTickerCount As Integer


Dim TickerColumn As Integer

TickerColumn = 1

Dim OpenColumn As Integer

OpenColumn = 3

Dim CloseColumn As Integer

CloseColumn = 6

Dim VolumeColumn As Integer

VolumeColumn = 7

Dim TickerTallyColumn As Integer

TickerTallyColumn = 9

Dim TotalColumn As Integer

TotalColumn = 12

Rem Create Ticker Array



Dim TickerArray(1 To 5000) As String

TickerArray(1) = " "

Dim OpenArray(1 To 5000) As Double

OpenArray(1) = 0

Dim CloseArray(1 To 5000) As Double

CloseArray(1) = 0

Dim Results As Range


For Each ws In Worksheets

    NewTickerCount = 2
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    TickerArray(NewTickerCount) = ws.Cells(NewTickerCount, TickerColumn).Value
    OpenArray(NewTickerCount) = ws.Cells(NewTickerCount, OpenColumn).Value
   
      
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("J2:J" & LastRow).NumberFormat = "0.000000000"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("L2:L" & LastRow).NumberFormat = "0"

    For tickerrow = 2 To LastRow

        If ws.Cells(tickerrow, TickerColumn).Value = TickerArray(NewTickerCount) Then
        
            ws.Cells(NewTickerCount, TickerTallyColumn).Value = TickerArray(NewTickerCount)
            ws.Cells(NewTickerCount, TotalColumn).Value = ws.Cells(NewTickerCount, TotalColumn).Value + ws.Cells(tickerrow, VolumeColumn).Value
        
        Else
            ws.Cells(NewTickerCount, TickerTallyColumn + 1).Value = ws.Cells(tickerrow - 1, CloseColumn).Value - OpenArray(NewTickerCount)
            
            If ws.Cells(NewTickerCount, TickerTallyColumn + 1).Value > 0 Then ws.Cells(NewTickerCount, TickerTallyColumn + 1).Interior.ColorIndex = 4
            If ws.Cells(NewTickerCount, TickerTallyColumn + 1).Value < 0 Then ws.Cells(NewTickerCount, TickerTallyColumn + 1).Interior.ColorIndex = 3
            If OpenArray(NewTickerCount) <> 0 Then ws.Cells(NewTickerCount, TickerTallyColumn + 2).Value = ws.Cells(NewTickerCount, TickerTallyColumn + 1).Value / OpenArray(NewTickerCount)
            
           
            NewTickerCount = NewTickerCount + 1
            TickerArray(NewTickerCount) = ws.Cells(tickerrow, TickerColumn).Value
            ws.Cells(NewTickerCount, TickerTallyColumn).Value = TickerArray(NewTickerCount)
            ws.Cells(NewTickerCount, TotalColumn).Value = ws.Cells(NewTickerCount, TotalColumn).Value + ws.Cells(tickerrow, VolumeColumn).Value
            OpenArray(NewTickerCount) = ws.Cells(tickerrow, OpenColumn).Value
      
        End If
      
       

    Next tickerrow
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("O2").Value = "Greatest % Increase"
    Set Results = ws.Range("K2:K" & NewTickerCount)
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q2").Value = Application.WorksheetFunction.Max(Results)
    ws.Range("P2").Value = ws.Range("I" & Application.Match(ws.Range("Q2").Value, Results, 0) + 1)
       
    
    ws.Range("O3").Value = "Greatest % Decrease"
    Set Results = ws.Range("K2:K" & NewTickerCount)
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q3").Value = Application.WorksheetFunction.Min(Results)
    ws.Range("P3").Value = ws.Range("I" & Application.Match(ws.Range("Q3").Value, Results, 0) + 1)
    
    ws.Range("O4").Value = "Greatest Total Volume"
    Set Results = ws.Range("L2:L" & NewTickerCount)
    ws.Range("Q4").Value = Application.WorksheetFunction.Max(Results)
    ws.Range("P4").Value = ws.Range("I" & Application.Match(ws.Range("Q4").Value, Results, 0) + 1)

Next ws



End Sub
