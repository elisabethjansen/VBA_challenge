# Code for Macro "Summary"

Sub Summary()

Dim i As Long
    Dim LastRow As Long
    Dim Ticker As String
    Dim Summary_Row As Long
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Summary_Row = 2
    Yearly_Change = 0
    Percent_Change = 0
    Total_Volume = 0
    Open_Price = 0
    Close_Price = 0
    

    LastRow = ws.UsedRange.Rows.Count
    Open_Price = Cells(2, 3).Value
    
    For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
            ws.Range("I" & Summary_Row).Value = Ticker
        Close_Price = ws.Cells(i, 6)
            
        Yearly_Change = Close_Price - Open_Price
            ws.Range("J" & Summary_Row).Value = Yearly_Change
    
        Percent_Change = (Yearly_Change / Open_Price)
                ws.Range("K" & Summary_Row).Value = Percent_Change
                    ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                    
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Row).Value = Total_Volume
                
        Summary_Row = Summary_Row + 1
        Yearly_Change = 0
        Open_Price = ws.Cells(i + 1, 3).Value
        Close_Price = 0
        Total_Volume = 0
    Else
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Row).Value = Total_Volume
    End If
Next i

 For i = 2 To LastRow
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    Dim Percent_Range As Range
    Dim Volume_Range As Range
    Dim Max_Percent As Double
    Dim Min_Percent As Double
    Dim Max_Volume As Double
    
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    Set Percent_Range = ws.Range("K2:K" & LastRow)
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(Percent_Range)
        ws.Cells(2, 17).NumberFormat = "0.00%"
        Max_Percent = ws.Cells(2, 17).Value
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(Percent_Range)
        ws.Cells(3, 17).NumberFormat = "0.00%"
        Min_Percent = Cells(3, 17).Value
    Set Volume_Range = ws.Range("L2:L" & LastRow)
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(Volume_Range)
        Max_Volume = Cells(4, 17).Value
        
    For i = 2 To LastRow
        If ws.Cells(i, 11).Value = Max_Percent Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value = Min_Percent Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value = Max_Volume Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        End If
    Next i
Next ws
End Sub