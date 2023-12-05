Attribute VB_Name = "Module1"
Sub Stock_Market():


    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        
        Dim ticker As String
        Dim vol As Double
                
        vol = ws.Cells(2, 7).Value
        
        Dim ticker_open As Single
        Dim ticker_close As Single
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim new_table_row As Integer
                
        new_table_row = 2
                
        ws.Columns("J:J").NumberFormat = "0.00"
        ws.Columns("K:K").Style = "Percent"
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").Style = "Percent"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("O:O").ColumnWidth = 20
        ws.Columns("Q:Q").EntireColumn.AutoFit

        Dim lastrow As Long
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
        ticker_open = ws.Cells(2, 3).Value
            
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    
                ticker = ws.Cells(i, 1).Value
                        
                ticker_close = ws.Cells(i, 6).Value
                yearly_change = ticker_close - ticker_open
                percent_change = yearly_change / ticker_open
            
                ws.Range("I" & new_table_row).Value = ticker
                ws.Range("J" & new_table_row).Value = yearly_change
                ws.Range("K" & new_table_row).Value = percent_change
                ws.Range("L" & new_table_row).Value = vol
                        
                If yearly_change > 0 Then
                    ws.Range("J" & new_table_row).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Range("J" & new_table_row).Interior.Color = RGB(255, 0, 0)
                End If
                
                new_table_row = new_table_row + 1
                vol = ws.Cells(i + 1, 7).Value
                ticker_open = ws.Cells(i + 1, 3).Value
                        
            Else
                    
                vol = vol + ws.Cells(i + 1, 7).Value
                    
            End If
                   
            
        Next i
 
        
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_total_volume As Double
            
        greatest_increase = WorksheetFunction.Max(ws.Range("K:K"))
        greatest_decrease = WorksheetFunction.Min(ws.Range("K:K"))
        greatest_total_volume = WorksheetFunction.Max(ws.Range("L:L"))
            
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
            
        ws.Range("Q2").Value = greatest_increase
        ws.Range("Q3").Value = greatest_decrease
        ws.Range("Q4").Value = greatest_total_volume

        For i = 2 To new_table_row - 1
            If ws.Cells(i, 11).Value = greatest_increase Then
                ws.Cells(i, 9).Copy
                ws.Range("P2").PasteSpecial xlPasteValues
            End If
                    
            If ws.Cells(i, 11).Value = greatest_decrease Then
                ws.Cells(i, 9).Copy
                ws.Range("P3").PasteSpecial xlPasteValues
            End If
                    
            If ws.Cells(i, 12).Value = greatest_total_volume Then
                ws.Cells(i, 9).Copy
                ws.Range("P4").PasteSpecial xlPasteValues
            End If

        Next i
                
    Next ws

End Sub
