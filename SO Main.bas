Attribute VB_Name = "Main"
Sub SO()


    'SO - Sale Order
    ' lr - last row without any filter, containing all data.
    ' lr2 - last row with a filter applied (closed=false).
    ' lastRow - last row in the "SO Data" worksheet, used for looping through rows to modify TECHNICALLY COMPLETED entries.
    ' lr4 - last row in the "Cumulative data" worksheet, used for applying filters on the copied data.
    ' lr5 - last row in the "Variance" worksheet, used for performing tolerance checks.

    Dim month_current As String
    month_current = InputBox("For which month the data is being performed and what RA stage? E.g., September Prelim")

    Dim per_current As String
    per_current = InputBox("What's the current period? Type in format dd-mm-yyyy")

    ' Set references to both the final workbook and source workbook
    Dim FinalBook As Workbook
    Set FinalBook = Workbooks("KW PAEN  " & month_current & " RA  SO  " & per_current & ".xlsm")
    
    Dim SourceBook As Workbook
    Set SourceBook = Workbooks("KW RA Working.xlsm")

    ' Activate the "SO Raw" sheet in SourceBook
    SourceBook.Worksheets("SO Raw").Activate
    
    Dim lr As Integer
    ' Find the last row of data in "SO Raw" sheet
    lr = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row
    
    ' Clear the data in "SO Data" in FinalBook
    Worksheets("SO Data ").Activate
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    Range("A4:AK4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Autofill the data range
    Range("A3:AK3").Select
    Selection.AutoFill Destination:=Range("A3:AK" & lr), Type:=xlFillDefault
    
    ' Clear "SO Raw" data in FinalBook
    FinalBook.Worksheets("SO Raw").Activate
    Range("A3:B3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("D3:AD3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Copy data from SourceBook "SO Raw" to FinalBook
    SourceBook.Worksheets("SO Raw").Activate
    Range("A3:B" & lr).Select
    Selection.Copy
    
    FinalBook.Worksheets("SO Raw").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Copy remaining data (columns C to AC) from SourceBook "SO Raw" to FinalBook
    SourceBook.Worksheets("SO Raw").Activate
    Range("C3:AC" & lr).Select
    Selection.Copy
    
    FinalBook.Worksheets("SO Raw").Activate
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Clear and autofill column C in "SO Raw"
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C" & lr), Type:=xlFillDefault

    ' Clearing data in "SO Data" in FinalBook
    FinalBook.Worksheets("SO Data").Activate
    Range("A4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Clear data in column E and columns H to AN
    Range("E4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("H4:AN4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Copy and filter data from "SO Data" in SourceBook
    SourceBook.Worksheets("SO Data ").Activate
    Range("A2").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$2:$AK$" & lr).AutoFilter Field:=8, Criteria1:="=RELEASED", Operator:=xlOr, Criteria2:="=TECHNICALLY COMPLETED"
    
    ' Find the last row with filtered data
    Dim lr2 As Integer
    lr2 = ActiveSheet.Range("$A$2:$A$" & lr).SpecialCells(xlCellTypeVisible).Count - 1
    
    ' Copy and paste visible cells from SourceBook to FinalBook
    ActiveSheet.Range("$A$2:$C$" & lr).SpecialCells(xlCellTypeVisible).Copy
    FinalBook.Worksheets("SO Data").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Copy column D data
    SourceBook.Worksheets("SO Data ").Activate
    ActiveSheet.Range("D2:D" & lr).SpecialCells(xlCellTypeVisible).Copy
    FinalBook.Worksheets("SO Data").Activate
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Copy remaining data from column E to AK in SourceBook to FinalBook
    SourceBook.Worksheets("SO Data ").Activate
    ActiveSheet.Range("E2:AK" & lr).SpecialCells(xlCellTypeVisible).Copy
    FinalBook.Worksheets("SO Data").Activate
    Range("H3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Clearing column D and autofill
    Range("D5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("D4").Select
    Selection.AutoFill Destination:=Range("D4:D" & lr2 + 3)
    
    ' Clearing columns F and G and autofill
    Range("F5:G5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("F4:G4").Select
    Selection.AutoFill Destination:=Range("F4:G" & lr2 + 3)

    ' Formatting columns A and C as General
    Range("A4:A" & lr).Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    Range("C4:C" & lr).Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    ' Apply filter to FinalBook and copy data to "TECO with OBL" worksheet
    Range("A3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AN$1132").AutoFilter Field:=11, Criteria1:=Array("Line Item Status", "TECHNICALLY COMPLETED", "="), Operator:=xlFilterValues
    
    ' Clear contents in "TECO with OBL" sheet
    FinalBook.Worksheets("TECO with OBL").Activate
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    Range("A4:AN4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Copy "SO Data" to "TECO with OBL"
    FinalBook.Worksheets("SO Data").Activate
    Range("A3:AN3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    FinalBook.Worksheets("TECO with OBL").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Set lastRow for the loop
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    FinalBook.Worksheets("SO Data").Activate
    Set wsData = Worksheets("SO Data")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the rows to update TECHNICALLY COMPLETED entries
    For i = 2 To lastRow
        If wsData.Cells(i, "K").Value = "TECHNICALLY COMPLETED" Then
            wsData.Cells(i, "AA").Value = 0
            wsData.Cells(i, "AB").Value = 0
            wsData.Cells(i, "AC").Value = 0
            wsData.Cells(i, "AD").Value = 0
        End If
    Next i
    
    ' Refresh pivot tables and copy data from pivot to periodic and cumulative data sheets
    FinalBook.Worksheets("Cumulative pivot").Visible = xlSheetVisible
    Worksheets("Periodic Pivot").Visible = xlSheetVisible
    Worksheets("Periodic Pivot").PivotTables("PivotTable1").RefreshTable
    Worksheets("Cumulative pivot").PivotTables("PivotTable1").RefreshTable
    
    ' Clean up Periodic Data and update it from the pivot
    FinalBook.Worksheets("Periodic data").Activate
    Range("A5:S5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    FinalBook.Worksheets("Periodic Pivot").Activate
    Range("A7:S7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    FinalBook.Worksheets("Periodic data").Activate
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Worksheets("Periodic Pivot").Visible = xlSheetHidden
    
    ' Clean up and update cumulative data
    FinalBook.Worksheets("Cumulative data").Activate
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    Range("A5:T5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    FinalBook.Worksheets("Cumulative pivot").Activate
    Range("A7:T7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    FinalBook.Worksheets("Cumulative data").Activate
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Worksheets("Cumulative Pivot").Visible = xlSheetHidden
    
    ' Apply filter to Cumulative data and copy to Variance
    FinalBook.Worksheets("Cumulative data").Activate
    Range("B4").Select
    Selection.AutoFilter
    ActiveSheet.Range("A4:T" & lr4).AutoFilter Field:=2, Criteria1:="RELEASED"
    ActiveSheet.Range("A4:T" & lr4).AutoFilter Field:=12, Criteria1:="<>0"
    FinalBook.Worksheets("Variance").Activate
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    Range("A4:I4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    ' Copy data from Cumulative Data to Variance
    FinalBook.Worksheets("Cumulative data").Activate
    ActiveSheet.Range("A4:E" & lr4).SpecialCells(xlCellTypeVisible).Copy
    FinalBook.Worksheets("Variance").Activate
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Copy additional columns to Variance
    FinalBook.Worksheets("Cumulative data").Activate
    ActiveSheet.Range("H4:I" & lr4).SpecialCells(xlCellTypeVisible).Copy
    FinalBook.Worksheets("Variance").Activate
    Range("F3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    FinalBook.Worksheets("Cumulative data").Activate
    ActiveSheet.Range("L4:M" & lr4).SpecialCells(xlCellTypeVisible).Copy
    FinalBook.Worksheets("Variance").Activate
    Range("H3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' AutoFill formulas in columns J and K of Variance sheet
    Dim wsData5 As Worksheet
    Dim lr5 As Integer
    FinalBook.Worksheets("Variance").Activate
    Set wsData5 = Worksheets("Variance")
    lr5 = wsData5.Cells(wsData5.Rows.Count, "A").End(xlUp).Row
    Range("J4:K4").Select
    Selection.AutoFill Destination:=Range("J4:K" & lr5), Type:=xlFillDefault

    ' Delete rows with values in column K that are below tolerance
    Dim tolerance As Double
    tolerance = 0.009
    Dim y As Long
    FinalBook.Worksheets("Variance").Activate
    For y = lr5 To 4 Step -1
        If Abs(wsData5.Cells(y, "K").Value) < tolerance Then
            wsData5.Rows(y).Delete
            DoEvents
        End If
    Next y

    ' Apply final sort and filter on Variance
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Variance").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Variance").AutoFilter.Sort.SortFields.Add2 Key:=Range("K3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Variance").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Refresh all data in the workbook
    ThisWorkbook.RefreshAll

    MsgBox "Done"

End Sub















