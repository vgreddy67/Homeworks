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
    
    
    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        prevTick = ws.Cells(2, 1).Value
        pos = 2
        stockVolume = 0
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
        For i = 2 To lastRow
            
            curTick = ws.Cells(i, 1).Value
        
            If curTick = prevTick Then
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            Else
               'adding ticker
                ws.Cells(pos, 9).Value = prevTick
                
               'adding volume for each ticker
                ws.Cells(pos, 10).Value = stockVolume
                ws.Cells(pos, 10).HorizontalAlignment = xlRight
                stockVolume = ws.Cells(i, 7).Value
                pos = pos + 1
            End If
            prevTick = curTick
        Next i
    Next ws
End Sub

