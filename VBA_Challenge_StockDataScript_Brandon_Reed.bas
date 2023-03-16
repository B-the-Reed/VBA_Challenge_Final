Attribute VB_Name = "Module2"
Sub StockDataAll()

For Each ws In Worksheets

Dim OpenPrice As Double
Dim ClosePrice As Double

Dim YearlyChange As Double
Dim PercentChange As Double

Dim TotalStockVolume As LongLong
Dim Ticker As String

Dim SummaryTable As LongLong

Dim lastrowticker As LongLong
Dim lastrowyearlychange As LongLong
Dim lastrowpercentchange As LongLong
Dim lastrowtotalstock As LongLong

Dim i As LongLong

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
lastrowticker = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastrowyearlychange = ws.Cells(Rows.Count, 10).End(xlUp).Row
lastrowpercentchange = ws.Cells(Rows.Count, 11).End(xlUp).Row

' Before I get into my for loop, I need to record the Open price for the first ticker

OpenPrice = ws.Range("C2").Value
SummaryTable = 2

For i = 2 To lastrowticker:

    ' If the next cell is different from the previous, then

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
        ' Record the Ticker symbol for i
        Ticker = ws.Cells(i, 1).Value
        
        ' Record the ClosePrice for i
        ClosePrice = ws.Cells(i, 6).Value
        
        ' Increment the TotalStock Volume by i
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        ' Calculate the Yearly Change (ClosePrice - OpenPrice)
        YearlyChange = ClosePrice - OpenPrice
        
        ' Calculate PercentChange (YearlyChange/OpenPrice)
        PercentChange = YearlyChange / OpenPrice
        
        ws.Cells(SummaryTable, 11).NumberFormat = "0.00%"
        
        ' Populate the Ticker symbol in table
        ws.Cells(SummaryTable, 9).Value = Ticker
        
        ' Populate YearlyChange in table
        ws.Cells(SummaryTable, 10).Value = YearlyChange
        
            If ws.Cells(SummaryTable, 10).Value > 0 Then
    
                ws.Cells(SummaryTable, 10).Interior.Color = vbGreen
    
            ElseIf ws.Cells(SummaryTable, 10).Value < 0 Then
        
                ws.Cells(SummaryTable, 10).Interior.Color = vbRed
    
            Else
    
                ws.Cells(SummaryTable, 10).Interior.ColorIndex = 0
        
            End If
        
        ' Populate PercentChange in table
        ws.Cells(SummaryTable, 11).Value = PercentChange
        
        'Change PercentChange to a percentage
        ' ws.Cells(SummaryTable, 11).Style = "Percent"
        
        ' Populate TotalStockVolume in table
        ws.Cells(SummaryTable, 12).Value = TotalStockVolume
            
        ' Shift SummaryTable down 1 row
        SummaryTable = SummaryTable + 1
        
        ' OpenPrice = i + 1
        OpenPrice = ws.Cells(i + 1, 3).Value
        
        ' TotalStockVolume = 0
        TotalStockVolume = 0


    ' Otherwise, you need to capture these pieces of info for each:
    Else
    
        ' Increment the TotalStockVolume by i
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    End If

Next i


' creating the bonus table now...

Dim GPIncTicker As String
Dim GPDecTicker As String
Dim GTotVolTicker As String

Dim PercentInc As Double
Dim PercentDec As Double
Dim TotalVolume As LongLong

Dim GreatestPercentInc As Double
Dim GreatestPercentDec As Double
Dim GreatestTotalVol As LongLong

lastrowpercentchange = ws.Cells(Rows.Count, 11).End(xlUp).Row
lastrowtotalstock = ws.Cells(Rows.Count, 12).End(xlUp).Row
GreatestPercentInc = 0
GreatestPercentDec = 0
GreatestTotalVol = 0

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

i = 2

For i = 2 To lastrowpercentchange

    If ws.Cells(i, 11).Value > GreatestPercentInc Then
    
        GreatestPercentInc = ws.Cells(i, 11).Value
        GPIncTicker = ws.Cells(i, 9).Value
    
    ElseIf ws.Cells(i, 11).Value < GreatestPercentDec Then
    
        GreatestPercentDec = ws.Cells(i, 11).Value
        GPDecTicker = ws.Cells(i, 9).Value
    
    End If

Next i

ws.Cells(2, 16).Value = GPIncTicker
ws.Cells(3, 16).Value = GPDecTicker
ws.Cells(2, 17).Value = GreatestPercentInc
ws.Cells(2, 17).Style = "Percent"
ws.Cells(3, 17).Value = GreatestPercentDec
ws.Cells(3, 17).Style = "Percent"

i = 2
GreatestTotalVol = 0

For i = 2 To lastrowtotalstock

    If ws.Cells(i, 12).Value > GreatestTotalVol Then
    
        GreatestTotalVol = ws.Cells(i, 12).Value
        GTotVolTicker = ws.Cells(i, 9).Value

    End If

Next i

ws.Cells(4, 16).Value = GTotVolTicker
ws.Cells(4, 17).Value = GreatestTotalVol

Next

End Sub
