Attribute VB_Name = "Module1"
Sub challenge()

Dim YrChange As Double
Dim sOpen As Variant
Dim sClose As Variant
Dim Counter As Integer
Dim LastRow As Long
Dim stockcounter As Integer
Dim sVolume As LongLong
Dim gPercentIn As Double
Dim gPercentDe As Double
Dim gVolume As LongLong
Dim gTick1 As String
Dim gTick2 As String
Dim gTick3 As String



sOpen = 0
sClose = 0
stockcounter = 0
sVolume = 0
gPercentIn = 0
gPercentDe = 0
gVolume = 0

For Each ws In Worksheets


Counter = 2

'Make Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To LastRow

    sVolume = sVolume + Cells(i, 7).Value
    
    sOpen = (Cells(i, 3).Value) + sOpen
    sClose = Cells(i, 6).Value + sClose
    stockcounter = stockcounter + 1
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        YrChange = sOpen - sClose
        'YrChange = (sOpen / stockcounter) - (sClose / stockcounter)
        Cells(Counter, 10).Value = YrChange
        
        'Set conditional formatting to cells
        If Cells(Counter, 10).Value >= 0 Then
            Cells(Counter, 10).Interior.ColorIndex = 4
        Else
            Cells(Counter, 10).Interior.ColorIndex = 3
        End If
    
        Cells(Counter, 9).Value = Cells(i, 1).Value
        Cells(Counter, 11).Value = 100 * (YrChange / (sOpen / stockcounter))
        If Cells(Counter, 11).Value >= 0 Then
            Cells(Counter, 11).Interior.ColorIndex = 4
        Else
            Cells(Counter, 11).Interior.ColorIndex = 3
        End If
        
        'Check if greatest
        If (100 * (YrChange / (sOpen / stockcounter))) > gPercentIn Then
        gPercentIn = (100 * (YrChange / (sOpen / stockcounter)))
        gTick1 = Cells(Counter, 1).Value
        End If
        If (100 * (YrChange / (sOpen / stockcounter))) < gPercentDe Then
        gPercentDe = (100 * (YrChange / (sOpen / stockcounter)))
        gTick2 = Cells(Counter, 1).Value
        End If
        If sVolume > gVolume Then
        gVolume = sVolume
        gTick3 = Cells(Counter, 1).Value
        End If
        
    
        Cells(Counter, 12).Value = sVolume
        
        
        Counter = Counter + 1
        YrChange = 0
        sOpen = 0
        sClose = 0
        stockcounter = 0
        sVolume = 0
        
        End If
    


Next i
'MsgBox ws.Name
 

Next ws

Cells(2, 16).Value = gTick1
Cells(3, 16).Value = gTick2
Cells(4, 16).Value = gTick3

MsgBox ("dunz")


End Sub



