VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockVol()
    
    Dim pos As Integer
    Dim prevTick As String
    Dim curTick As String
    Dim openValue As Double
    Dim closeValue As Double
    Dim yearlyChange As Double
    
    
    
    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        prevTick = ws.Cells(2, 1).Value
        pos = 2
        openValue = ws.Cells(2, 3).Value
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        For i = 2 To lastRow
            
            curTick = ws.Cells(i, 1).Value
        
            If curTick = prevTick Then
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            Else
               'adding ticker
                ws.Cells(pos, 9).Value = prevTick
                
              'Yearly change
                closeValue = ws.Cells(i - 1, 6).Value
                
                'Formatting the cells
                yearlyChange = closeValue - openValue
                ws.Cells(pos, 10).NumberFormat = "0.000000000"
                ws.Cells(pos, 10).Value = yearlyChange
                
                'Adding green to cell for positive change and red for negative change
                If yearlyChange < 0 Then
                    ws.Cells(pos, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(pos, 10).Interior.ColorIndex = 4
                End If
                
              'percentageChange
                If openValue <> 0 Then
                    percentageChange = (yearlyChange / openValue)
                    'percentageChange = (yearlyChange * 100) / openValue
                    ws.Cells(pos, 11).NumberFormat = "0.00%"
                    ws.Cells(pos, 11).Value = percentageChange
                Else
                    percentageChange = 0
                    ws.Cells(pos, 11).NumberFormat = "0.00%"
                    ws.Cells(pos, 11).Value = percentageChange
                End If
                
                'adding volume for each ticker
                ws.Cells(pos, 12).Value = stockVolume
                ws.Cells(pos, 12).HorizontalAlignment = xlRight
                openValue = ws.Cells(i, 3).Value
                stockVolume = ws.Cells(i, 7).Value
            
                pos = pos + 1
                
            End If
            prevTick = curTick
        Next i
    Next ws
End Sub
