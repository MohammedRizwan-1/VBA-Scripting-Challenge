
'Stock_Testing_VBA_Analsysis

    Sub Stock_Testing_Data():

        '1.0 Defining Given Variables in Sheet

            Dim WorksheetName as String
            Dim Ticker As String
            Dim Stock_Volume As LongLong
            Dim Yearly_Open As Double
            Dim Yearly_Close As Double
            Dim YR_Change As Double
            Dim Percent_Change As Double
            Dim Summary_Table_Row As Double
            Dim Start_Row As Double
            
        '1.0.2 Worksheet Setup
        For Each ws in Worksheets

            '1.1 Set Headers For Results Columns
                ws.Cells(1, 10).Value = "Ticker"
                ws.Cells(1, 11).Value = "Yearly Change"
                ws.Cells(1, 12).Value = "Percentage Change"
                ws.Cells(1, 13).Value = "Total Stock Volume"

            '1.2 Integers for loop
                Summary_Table_Row = 2
                Stock_Volume = 0
                Start_Row = 2
            
            '1.3 Define Number of Rows
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Next ws
End Sub
        