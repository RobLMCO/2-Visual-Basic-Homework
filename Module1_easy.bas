Attribute VB_Name = "Module1"
Sub TickerTally()

Rem Declare Variables

Dim WorksheetName As String

Dim NewTickerCount As Integer



Dim TickerColumn As Integer

TickerColumn = 1

Dim VolumeColumn As Integer

VolumeColumn = 7

Dim TotalColumn As Integer

TotalColumn = 10

Rem Create Ticker Array



Dim TickerArray(1 To 5000) As String

TickerArray(1) = " "

Dim TickerTotal As Integer

TickerTotal = 0

Rem WorksheetName = ws.Name

For Each ws In Worksheets

    NewTickerCount = 2
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    TickerArray(NewTickerCount) = ws.Cells(NewTickerCount, TickerColumn).Value
    
    ws.Cells(1, TotalColumn - 1).Value = "Ticker"
    ws.Cells(1, TotalColumn).Value = "Total Stock Volume"

    For tickerrow = 2 To LastRow

        If ws.Cells(tickerrow, TickerColumn).Value = TickerArray(NewTickerCount) Then
        
            ws.Cells(NewTickerCount, TotalColumn - 1).Value = TickerArray(NewTickerCount)
            ws.Cells(NewTickerCount, TotalColumn).Value = ws.Cells(NewTickerCount, TotalColumn).Value + ws.Cells(tickerrow, VolumeColumn).Value
        
        Else
            NewTickerCount = NewTickerCount + 1
            TickerArray(NewTickerCount) = ws.Cells(tickerrow, TickerColumn).Value
            ws.Cells(NewTickerCount, TotalColumn - 1).Value = TickerArray(NewTickerCount)
            ws.Cells(NewTickerCount, TotalColumn).Value = ws.Cells(NewTickerCount, TotalColumn).Value + ws.Cells(tickerrow, VolumeColumn).Value
      
        End If
      
       Rem ws.Cells(1, 1).Value = "State"

    Next tickerrow

Next ws


End Sub
