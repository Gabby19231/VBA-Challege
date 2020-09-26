# VBA-Challenge

VBA scripting was completed to analyze stock market data.

The script was created to look through all of the stocks for one year. Along with the output of that information.

some code I was able to create is attached below. This is used to start finding the tickers and yearly changes.

Sub ticker()

Dim ticker As String

Dim ticker As Integer

For Each ws In Worksheets
     ws.Activate
     
Starting_Value = 2
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 ws.Range("A1").EntireColumn.Insert
 
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    
    number_tickers = 0
    yearly_change = 0

       LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        For H = 3 To LastColumn
            TickerHeader = ws.Cells(1, H).Value
            YearSplit = Split(YearHeader, " ")
            
            ws.Cells(1, H).Value = TickerSplit(3)

        Next H

        For H = 2 To LastRow

            For G = 3 To LastColumn

                ws.Cells(H).Style = "Tickers"

            Next G

        Next H

    Next ws

    MsgBox ("Tickers Complete")
    MsgBox ("Yearly Chanages Complete")


End Sub
