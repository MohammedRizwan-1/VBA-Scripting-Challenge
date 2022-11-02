
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

            For  i = 2 To LastRow

                'Check if on same Stock Ticker, If not then
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                
                
                'Set Ticker Value and Print It in Column
                    Ticker = ws.Cells(i, 1).Value
                    ws.Range("J" & Summary_Table_Row).Value = Ticker
                
                'Add to Stock Volume and Print it in Column
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                    ws.Range("M" & Summary_Table_Row).Value = Stock_Volume
                
                'Yearly Change
                    YR_Change = ws.Cells(i, 6).Value - ws.Cells(Start_Row, 3)
                    ws.Range("K" & Summary_Table_Row).Value = YR_Change

                'Formatting Cell Colours 
                    If ws.Range("K"& Summary_Table_Row).Value >=0 Then
                        ws.Range("K"& Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf ws.Range("K"& Summary_Table_Row).Value <=0 Then
                        ws.Range("K"& Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                'Set Percetage Change Values and Print it in Column
                    Percentage_Change = YR_Change / ws.Cells(Start_Row, 3)
                    
                    ws.Range("L" & Summary_Table_Row).Value = Percentage_Change
                    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"


                'Increase Summary Table Row by 1
                    Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset Stock Volume
                    Stock_Volume = 0
                
                    Start_Row = (i + 1)

                'If Cell following Row is the same ticker symbol then
                Else
                    ' Add To Stock Volume
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                End If

        Next ws
End Sub
        