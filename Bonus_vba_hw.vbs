Option Explicit
Sub TickerCalculationBonus():

'Data
'Col 1 - Ticker
'Col 2 - Date
'Col 3 - Open
'Col 4 - High
'Col 5 - Low
'Col 6 - Close
'Col 7 - Volume
'Col 8 - Volume in 000s

'Output
'Col 9 - Ticker
'Col 10- - Yearly Change
'Col 11 - Percentage Change
'Col 12 - Total Stock Volume
    

Dim dataRow As Long
Dim outputRow As Long
Dim sheetNum As Long

'I know this is first ticker, first row
'therefore, save the open price
Dim openPrice As Double
Dim totalStockVolume As Double
Dim closePrice As Double

For sheetNum = 1 To Worksheets.Count

    Worksheets(sheetNum).Activate
    
    outputRow = 2
    
    openPrice = ActiveSheet.Range("C2").Value
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Percentage Change"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("Q1").Select
        ActiveCell.FormulaR1C1 = "Value"
    Range("O2").Select
        ActiveCell.FormulaR1C1 = "Greatest % increase"
    Range("O3").Select
        ActiveCell.FormulaR1C1 = "Greatest % decrease"
    Range("O4").Select
        ActiveCell.FormulaR1C1 = "Greatest total volume"

    Columns("J:J").Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=2
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -0.01
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 5287936
        .TintAndShade = 0
    End With
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    
    'Start loop at A2
    For dataRow = 2 To ActiveSheet.Range("A2").End(xlDown).Row
        If ActiveSheet.Cells(dataRow, 1).Value <> ActiveSheet.Cells(dataRow + 1, 1).Value Then
            'Now at the edge
            'add whatever is in Col G to the total stock volume counter
            totalStockVolume = totalStockVolume + ActiveSheet.Cells(dataRow, 7).Value / 1000
            'grab the closing proce from Col f
            closePrice = ActiveSheet.Cells(dataRow, 6).Value
            'Now calculate yearly change as close_price - open_price
            'Calculate yearly percenta change as close_price - open_price / open_price
            'Since there might be a division by 0, put in check to make sure that the denominator is not 0
            'Copy over the value in Col A to Col I
            'Then dump the yearly change, percent change and total stock volume into j,k,l
            'Percentage Change
            If openPrice = 0 Then
                ActiveSheet.Cells(outputRow, 11).Value = "NaN"
            Else
                ActiveSheet.Cells(outputRow, 11).Value = (closePrice - openPrice) / openPrice
            End If
            'Yearly Change
            ActiveSheet.Cells(outputRow, 10).Value = closePrice - openPrice

            'Total Stock Volume
            ActiveSheet.Cells(outputRow, 12).Value = totalStockVolume * 1000

            'Ticker
            ActiveSheet.Cells(outputRow, 9).Value = ActiveSheet.Cells(dataRow, 1).Value

            'Add 1 to the row counter for the output table
            outputRow = outputRow + 1
            'Then update the new open price to be the open price of the new row
            totalStockVolume = 0
            openPrice = ActiveSheet.Cells(dataRow + 1, 3).Value
        Else
            'If it's not the edge, then
            'don't change the open value
            'add whatever is in Col G to the total stock counter
            totalStockVolume = totalStockVolume + ActiveSheet.Cells(dataRow, 7).Value / 1000
        End If
    Next dataRow
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[-6]:R[1048574]C[-6])"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[-1]C[-6]:R[1048573]C[-6])"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-2]C[-5]:R[1048572]C[-5])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[1],RC[-5]:R[1048574]C[-5],RC[-7]:R[1048574]C[-7])"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[1],R[-1]C[-5]:R[1048573]C[-5],R[-1]C[-7]:R[1048573]C[-7])"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = _
        "=XLOOKUP(RC[1],R[-2]C[-4]:R[1048572]C[-4],R[-2]C[-7]:R[1048572]C[-7])"
Next sheetNum

End Sub