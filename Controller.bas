Attribute VB_Name = "Controller"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxRow(1)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function


Sub PO_Template()
'
' PO_Template Macro
'

    Dim file As String
    file = Range("B6").Value
    
    Dim fn As String
    fn = GetFilenameFromPath(file)
    
    Dim open_po As String
    open_po = Range("B2").Value
    
    Workbooks.Open fileName:=file
    Workbooks(fn).Activate
    
    Range(Columns(1), Columns(findCellInColumn(1, "Material Number") - 1)).Select
    Selection.EntireColumn.Hidden = True
    
    Range(Columns(findCellInColumn(1, "Material Number") + 1), Columns(findCellInColumn(1, "Product line") - 1)).Select
    Selection.EntireColumn.Hidden = True
    
    Range(Columns(findCellInColumn(1, "Product line") + 1), Columns(getMaxCol(1) - 1)).Select
    Selection.EntireColumn.Hidden = True
    
    Columns(findCellInColumn(1, "Material Number")).EntireColumn.AutoFit
    Cells(1, findCellInColumn(1, "Material Number")).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns(findCellInColumn(1, "Product line")).EntireColumn.AutoFit
    Cells(1, findCellInColumn(1, "Product line")).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    x = getMaxCol(1)
    
    Cells(1, x).Value = "Plant"
    Cells(1, x + 1).Value = "Family"
    Cells(1, x + 2).Value = "PIC"
    Cells(1, x + 3).Value = "Remarks"
    
    Range(Cells(1, x), Cells(1, x + 3)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    
    Workbooks.Open fileName:=open_po
    Workbooks(GetFilenameFromPath(open_po)).Activate
    ActiveWorkbook.Worksheets("Workbook").Activate
    mn_po = findCellInColumn(1, "Material Number")
    pl_po = findCellInColumn(1, "Plant")
    x_po = getMaxCol(1)
    
    Workbooks(fn).Activate
        
    n = x - findCellInColumn(1, "Material Number")
        
    Cells(2, x).Activate
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-" + CStr(n) + "],'[" + GetFilenameFromPath(open_po) + "]Workbook'!C" + CStr(mn_po) + ":C" + CStr(pl_po) + "," + CStr(x_po - 2 - mn_po) + ",0)"
    Selection.AutoFill Destination:=Range(Cells(2, x), Cells(getMaxRow(1), x))
    
    Cells(2, x + 1).Activate
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-" + CStr(n + 1) + "],'[" + GetFilenameFromPath(open_po) + "]Workbook'!C" + CStr(mn_po) + ":C" + CStr(pl_po + 1) + "," + CStr(x_po - 1 - mn_po) + ",0)"
    Selection.AutoFill Destination:=Range(Cells(2, x + 1), Cells(getMaxRow(1), x + 1))
    
    Cells(2, x + 2).Activate
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-" + CStr(n + 2) + "],'[" + GetFilenameFromPath(open_po) + "]Workbook'!C" + CStr(mn_po) + ":C" + CStr(pl_po + 2) + "," + CStr(x_po - mn_po) + ",0)"
    Selection.AutoFill Destination:=Range(Cells(2, x + 2), Cells(getMaxRow(1), x + 2))
    
    Range(Columns(findCellInColumn(1, "Material Number")), Columns(x + 3)).Select
  '  Selection.AutoFilter Field:=(n + 1), Criteria1:="PEL"
 '   ActiveSheet.Range(Cells(1, findCellInColumn(1, "Material Number")), Cells(getMaxRow(1), (x + 3))).AutoFilter Field:=(n + 1), Criteria1:="PEL"
    
    Range(Columns(x), Columns(x + 3)).Select
    Selection.Columns.AutoFit
    
    BreakLinks
End Sub

Sub PO_Pivot(fn As String, ob As Integer)
'
' PO_Pivot Macro
'

'
    Dim file As String
    file = Range("B6").Value
    
'    Dim fn As String
 '   fn = GetFilenameFromPath(file)
    Workbooks(fn).Activate
    
    Dim corner As String
    max_row = getMaxRow(1)
    x = getMaxCol(1)
    Range(Cells(1, 1), Cells(max_row, x)).Select
    Dim source As String
    source = "Workbook!R1C1:R" + CStr(max_row) + "C" + CStr(x)
    Sheets.Add.Name = "PivotTable"
'    MsgBox source
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6).CreatePivotTable TableDestination:= _
        "PivotTable!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("PivotTable").Select
    
'    Columns("A:DI").Select
'    Sheets.Add
 '   ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        "Workbook!R1C1:R1048576C113", Version:=7).CreatePivotTable TableDestination _
 '       :="Sheet3!R3C1", TableName:="PivotTable4", DefaultVersion:=7
  '  Sheets("Sheet3").Select
   ' Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
            .RefreshOnFileOpen = False
            .MissingItemsLimit = xlMissingItemsDefault
        End With
        
    If ob = 1 Then
        ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material Number")
            .Orientation = xlPageField
            .Position = 1
        End With
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Plant")
            .Orientation = xlRowField
            .Position = 1
        End With
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
            "PivotTable1").PivotFields("Open Quantity"), "Count of Open Quantity", xlCount
    ElseIf ob = 2 Then
        ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Plant")
            .Orientation = xlPageField
            .Position = 1
        End With
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("ULO Ageing Category")
            .Orientation = xlRowField
            .Position = 1
        End With
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
            "PivotTable1").PivotFields("Open Quantity"), "Count of Open Quantity", xlCount
        ActiveSheet.PivotTables("PivotTable1").PivotFields("Plant").CurrentPage = "PEL"
    End If
End Sub

Sub BreakLinks()
    Dim Links As Variant
    Links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    For i = 1 To UBound(Links)
    ActiveWorkbook.BreakLink _
        Name:=Links(i), _
        Type:=xlLinkTypeExcelLinks
    Next i
End Sub
