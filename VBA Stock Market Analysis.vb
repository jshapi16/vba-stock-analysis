Option Explicit
Sub VBAHWTest()

Dim openPrice As Double
openPrice = Range("C2").Value
Dim dataRow As Long
Dim outputRow As Long
Dim totalStock As Long
Dim totalStockVolume As Double
Dim closePrice As Double
Dim sheetNum As Long
Dim ws As Worksheet
Dim lastRow As Double
Dim rowData As Long

For sheetNum = 1 To Worksheets.Count
    Worksheets(sheetNum).Activate
    outputRow = 2
'Insert column for ticker, yr_chg, percent_chg, total volume
    ActiveSheet.Range("I1").Value = "Ticker"
    ActiveSheet.Range("J1").Value = "Year Change"
    ActiveSheet.Range("K1").Value = "Percent Change"
    ActiveSheet.Range("L1").Value = "Total Volume"

For dataRow = 2 To ActiveSheet.Range("A2").End(xlDown).Row
    If ActiveSheet.Cells(dataRow, 1).Value <> ActiveSheet.Cells(dataRow + 1, 1).Value Then
        'Total Stock Volume
        totalStockVolume = totalStockVolume + ActiveSheet.Cells(dataRow, 7).Value
        closePrice = ActiveSheet.Cells(dataRow, 6).Value
        'Calculate percent change
        If openPrice = 0 Then
            ActiveSheet.Cells(outputRow, 11).Value = "NaN"
        Else
            ActiveSheet.Cells(outputRow, 11).Value = (closePrice - openPrice) / openPrice
        End If
        'Yearly Change
        ActiveSheet.Cells(outputRow, 10).Value = closePrice - openPrice
        'Total Stock Volume
         ActiveSheet.Cells(outputRow, 12).Value = totalStockVolume
        'Ticker
        ActiveSheet.Cells(outputRow, 9).Value = ActiveSheet.Cells(dataRow, 1).Value
        'Add 1 to the row counter for the output table
        outputRow = outputRow + 1
        'Update new open price
        totalStockVolume = 0
        openPrice = ActiveSheet.Cells(dataRow + 1, 3).Value
    Else
        totalStockVolume = totalStockVolume + ActiveSheet.Cells(dataRow, 7).Value
    End If
Next dataRow

lastRow = ActiveSheet.Cells(Rows.Count, 11).End(xlUp).Row
For rowData = 2 To lastRow
    If ActiveSheet.Cells(rowData, 11).Value < 0 Then
        ActiveSheet.Cells(rowData, 11).Interior.ColorIndex = 3
    ElseIf ActiveSheet.Cells(rowData, 11).Value > 0 Then
        ActiveSheet.Cells(rowData, 11).Interior.ColorIndex = 4
    Else ActiveSheet.Cells(rowData, 11).Interior.ColorIndex = 2
    End If

Next rowData

Next sheetNum

End Sub
