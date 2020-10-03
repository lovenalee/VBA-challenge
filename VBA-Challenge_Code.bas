Attribute VB_Name = "Module1"
'Questions:
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.

Sub TickerSymbol():
    
    Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim TickerSymbol As String
            Dim Summary_Table_Row As Integer
        
                Summary_Table_Row = 2
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                For i = 2 To LastRow
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        TickerSymbol = ws.Cells(i, 1).Value
                        ws.Range("K" & Summary_Table_Row).Value = TickerSymbol
                        Summary_Table_Row = Summary_Table_Row + 1
                    End If
                Next i
    
        Next ws
    

End Sub
Sub YearlyChange():
    
    Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim YearlyChange As Double
            Dim Summary_Table_Row As Integer
            Dim RowCount As LongLong
        
                Summary_Table_Row = 2
                RowCount = 0
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                For i = 2 To LastRow
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        YearlyChange = (ws.Cells(i, 6).Value - ws.Cells(i - RowCount, 3).Value)
                        ws.Range("L" & Summary_Table_Row).Value = YearlyChange
                        Summary_Table_Row = Summary_Table_Row + 1
                        RowCount = 0
                    Else
                        RowCount = 1 + RowCount
                    End If

                Next i
    
        Next ws
    
End Sub

Sub ColorChange():

    Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow
                If ws.Cells(i, 12).Value >= 0 Then
                    ws.Cells(i, 12).Interior.ColorIndex = 4
                Else
                    ws.Cells(i, 12).Interior.ColorIndex = 3
    
                End If
            Next i
        Next ws
                    
End Sub

                    
Sub PercentChange()

    Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim PercentChange As Double
            Dim Summary_Table_Row As Integer
            Dim RowCount As LongLong

        
                Summary_Table_Row = 2
                RowCount = 0
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                For i = 2 To LastRow
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        PercentChange = (ws.Cells(i, 6).Value - ws.Cells(i - RowCount, 3)) / IIf(ws.Cells(i - RowCount, 3) = 0, 1, ws.Cells(i - RowCount, 3).Value)
                        ws.Range("M" & Summary_Table_Row).Value = PercentChange
                        ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"
                        Summary_Table_Row = Summary_Table_Row + 1
                        RowCount = 0
                    Else
                        RowCount = 1 + RowCount
                    End If
                    
                    
                Next i
    
        Next ws
    
End Sub
Sub TotalStockVolume()

    Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim TotalStockVolume As LongLong
            Dim Summary_Table_Row As Integer

                TotalStockVolume = 0
                Summary_Table_Row = 2
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                For i = 2 To LastRow
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        ws.Range("N" & Summary_Table_Row).Value = TotalStockVolume
                        Summary_Table_Row = Summary_Table_Row + 1
                        TotalStockVolume = 0
                    Else
                        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                    End If
                    
                    
                Next i
    
        Next ws
    
End Sub

Sub Greatest_Increase():

Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim Max As Double
            Dim MaxTicket As String
            Dim Min As Double
            Dim MinTicket As String
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Max = 0
    
            For i = 2 To LastRow
                For Each Cell In ws.Cells(i, 13)
                    If ws.Cells(i, 13).Value > Max Then
                        Max = ws.Cells(i, 13).Value
                        MaxTicket = ws.Cells(i, 11).Value
                    End If
                Next
            Next
                
            ws.Cells(2, 18).Value = Max
            ws.Cells(2, 18).NumberFormat = "0.00%"
            ws.Cells(2, 17).Value = MaxTicket
        Next ws
                    
End Sub

Sub Greatest_Decrease():

Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim Min As Double
            Dim MinTicket As String
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Min = 0
    
            For i = 2 To LastRow
                For Each Cell In ws.Cells(i, 13)
                    If ws.Cells(i, 13).Value < Min Then
                        Min = ws.Cells(i, 13).Value
                        MinTicket = ws.Cells(i, 11).Value
                    End If
                Next
            Next
                
            ws.Cells(3, 18).Value = Min
            ws.Cells(3, 18).NumberFormat = "0.00%"
            ws.Cells(3, 17).Value = MinTicket
        Next ws
                    
End Sub

Sub Greatest_Total_Volume():

Dim ws As Worksheet
        For Each ws In Worksheets
        
            Dim LastRow As Long
            Dim MaxVolume As Double
            Dim MaxVolumeTicket As String
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            MaxVolume = 0
    
            For i = 2 To LastRow
                For Each Cell In ws.Cells(i, 14)
                    If ws.Cells(i, 14).Value > MaxVolume Then
                        MaxVolume = ws.Cells(i, 14).Value
                        MaxVolumeTicket = ws.Cells(i, 11).Value
                    End If
                Next
            Next
                
            ws.Cells(4, 18).Value = MaxVolume
            ws.Cells(4, 17).Value = MaxVolumeTicket
        Next ws
                    
End Sub
Sub SummaryTableTittle():


    Dim ws As Worksheet
        For Each ws In Worksheets

            ws.Cells(1, 11).Value = "Ticker Symbol"
            ws.Cells(1, 12).Value = "Yearly Change"
            ws.Cells(1, 13).Value = "Percent Change"
            ws.Cells(1, 14).Value = "Total Stock Volume"
            ws.Cells(1, 17).Value = "Ticker"
            ws.Cells(1, 18).Value = "Value"
            ws.Cells(2, 16).Value = "Greatest % Increase"
            ws.Cells(3, 16).Value = "Greatest % Decrease"
            ws.Cells(4, 16).Value = "Greatest Total Volume"
    
        Next ws
            
End Sub



