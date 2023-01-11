
Sub TickerSymbol()

Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
'Determine the Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

' Assigning  Column Values

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Percent Change"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Total Stock Volume"
        

'Create Variable to hold Value
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim TickerName As String
Dim Percent As Double

Dim Volume As Double
Volume = 0

Dim Row As Double
Row = 2

Dim Column As Integer
Column = 1
Dim i As Long
        
        'Open Price Calculation
        
        
        OpenPrice = Cells(2, Column + 2).Value
         
        ' Loop for Ticker
        
        For i = 2 To LastRow
        
         ' Comparison using <> symbol
         
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                ' Get Details for Ticker
                
                
                TickerName = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = TickerName
                
                'Close Price
                ClosePrice = Cells(i, Column + 5).Value
                
                'Yearly Change
                YearlyChange = ClosePrice - OpenPrice
                Cells(Row, Column + 9).Value = YearlyChange
                
                'Percent Change
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    Percent = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    Percent = 1
                Else
                    Percent = YearlyChange / OpenPrice
                    Cells(Row, Column + 10).Value = Percent
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                
                'Total Volume
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                'Summar Table Row
                Row = Row + 1
                
                ' reset the Open Price
                OpenPrice = Cells(i + 1, Column + 2)
                
                ' reset the Volumn Total
                Volume = 0
            
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
            
            
            
        Next i
        
        
        
        ' Yearly Chnage
        LastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' Set the Cell Colors
        
        For j = 2 To LastRow
            If (Cells(j, Column + 10).Value > 0 Or Cells(j, Column + 10).Value = 0) Then
                Cells(j, Column + 10).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 10).Value < 0 Then
                Cells(j, Column + 10).Interior.ColorIndex = 3
            End If
        Next j
        
        'Last Part Greatest % Increase, % Decrease, and Total Volume
        
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        'Greatest Value Amongst Each Row
        
        For Z = 2 To LastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & LastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & LastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & LastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
             
        
    Next WS

End Sub
