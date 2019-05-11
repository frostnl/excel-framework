Option Explicit
Private pTable As ListObject
Private pSheet As Worksheet
Private pPrimaryKey As String
Private pLastRow As Long


'----------------------------------------------------------------------------------------------------
'*
'* InitialiseSheet
'*
'----------------------------------------------------------------------------------------------------
Function InitialiseSheet(sheet, pTableName, Optional pKey As String = "")
  pPrimaryKey = pKey
  
  pLastRow = 0
  Set pSheet = sheet
  Set pTable = pSheet.ListObjects(pTableName)
End Function



'----------------------------------------------------------------------------------------------------
'*
'* PrimaryKey
'*
'----------------------------------------------------------------------------------------------------
Public Property Get primaryKey() As String
  primaryKey = pPrimaryKey
End Property


Public Property Let LastRow(var As Long)
    pLastRow = var
End Property

Public Property Get LastRow() As Long
    If pLastRow = 0 Then
        LastRow = Me.sheet.Rows.count
    Else
        LastRow = pLastRow
    End If
End Property


'----------------------------------------------------------------------------------------------------
'*
'* Sheet
'*
'----------------------------------------------------------------------------------------------------
Function sheet() As Worksheet
  Set sheet = pSheet
End Function



'----------------------------------------------------------------------------------------------------
'*
'* HeaderVariantArray
'*
'----------------------------------------------------------------------------------------------------
Function HeaderVariantArray(Optional RemoveSpace As Boolean = False) As Variant
  Dim headers As Variant
  Dim i As Integer
  
  headers = pTable.HeaderRowRange
  For i = LBound(headers, 2) To UBound(headers, 2)
    headers(1, i) = Replace(headers(1, i), " ", "")
  Next i
    
  HeaderVariantArray = headers
End Function



'----------------------------------------------------------------------------------------------------
'*
'* RowVariantArray
'*
'----------------------------------------------------------------------------------------------------
Function RowVariantArray(tableRow) As Variant
  Dim data As Variant
  If tableRow <= DataRowCount Then
    data = pSheet.Range(pSheet.Cells(tableRow + DataStartCell.row - 1, DataStartCell.column), pSheet.Cells(tableRow + DataStartCell.row - 1, DataStartCell.column + TableColumnCount - 1))
  Else
    data = Nothing
  End If

  RowVariantArray = data
End Function



'----------------------------------------------------------------------------------------------------
'*
'* HeaderObject
'*
'----------------------------------------------------------------------------------------------------
Function HeaderObject() As Scripting.Dictionary
    Dim headerArray As Variant
    Dim dataArray As Variant
    Dim field As String
    headerArray = HeaderVariantArray

    'get header array
    Dim header As Scripting.Dictionary
    Set header = New Scripting.Dictionary
    For j = LBound(headerArray, 2) To UBound(headerArray, 2)
        field = headerArray(1, j)
        header.Add field, j
    Next j

    Set HeaderObject = header
End Function

Function GetColumnIndex(columnName As String) As Integer
    Dim column As Integer
    Dim j As Integer
    Dim headerArray As Variant
    headerArray = HeaderVariantArray
    column = 0
    For j = LBound(headerArray, 2) To UBound(headerArray, 2)
        If headerArray(1, j) = columnName Then
            column = j
        End If
    Next j
    GetColumnIndex = column
End Function


'----------------------------------------------------------------------------------------------------
'*
'* ClearTable
'*
'----------------------------------------------------------------------------------------------------
Function ClearTable()
On Error Resume Next
  pTable.DataBodyRange.ClearContents
End Function



'----------------------------------------------------------------------------------------------------
'*
'* DataStartCell
'*
'----------------------------------------------------------------------------------------------------
Function DataStartCell() As Range
   Set DataStartCell = pTable.DataBodyRange.Cells(1, 1)
End Function



'----------------------------------------------------------------------------------------------------
'*
'* TableStartCell
'*
'----------------------------------------------------------------------------------------------------
Function TableStartCell() As Range
  Set TableStartCell = pTable.HeaderRowRange.Cells(1, 1)
End Function



'----------------------------------------------------------------------------------------------------
'*
'* TableColumnCount
'*
'----------------------------------------------------------------------------------------------------
Function TableColumnCount() As Long
  TableColumnCount = pTable.ListColumns.count
End Function



'----------------------------------------------------------------------------------------------------
'*
'* DataRowCount
'*
'----------------------------------------------------------------------------------------------------
Function DataRowCount() As Long
  DataRowCount = pTable.ListRows.count
End Function



'----------------------------------------------------------------------------------------------------
'*
'* TableColumnStart
'*
'----------------------------------------------------------------------------------------------------
Function TableColumnStart() As Long
  TableColumnStart = pTable.HeaderRowRange.Cells(1, 1).column
End Function



'----------------------------------------------------------------------------------------------------
'*
'* ResizeTable
'*
'----------------------------------------------------------------------------------------------------
Function resizeTable(Optional dataRows As Long = -1)
  Dim Range As Range
  
    If dataRows <> -1 Then
      dataRows = dataRows + 1 'need to include header row
    Else
      dataRows = LastRowColumn(pSheet, "Row") - TableStartCell.row + 1
    End If
  
  
    If dataRows < 2 Then
      dataRows = 2
    End If
  
  Set Range = TableStartCell.Resize(dataRows, TableColumnCount)
  pTable.Resize Range

End Function



'----------------------------------------------------------------------------------------------------
'*
'* LastRowColumn
'*
'----------------------------------------------------------------------------------------------------
Function LastRowColumn(sht As Worksheet, RowColumn As String) As Long
'PURPOSE: Function To Return the Last Row Or Column Number In the Active Spreadsheet
'INPUT: "R" or "C" to determine which direction to search
Dim rc As Long

'clear filters
Me.ClearFilters

Select Case LCase(Left(RowColumn, 1)) 'If they put in 'row' or column instead of 'r' or 'c'.
  Case "c"
    LastRowColumn = sht.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious).column
  Case "r"
    LastRowColumn = sht.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious).row
  Case Else
    LastRowColumn = 1
End Select

End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchAsObjectDictionary
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllAsObjectDictionary(Optional pKey = "", Optional ignoreDuplicates As Boolean = False, Optional ExcludeFilteredRows As Boolean = False) As Scripting.Dictionary
    Dim arr As Variant
    Dim field As String
    Dim obj As Scripting.Dictionary
    Dim j As Long
    Dim i As Long
    Dim tempData As Variant
    Dim rng As Range
    

    Dim dataArray As Variant
    

    Dim numRows As Long
    Dim firstRow As Range
    Dim currentIndex As Long
    currentIndex = 1
    
    Dim row As Long
    Dim col As Long
    Dim area As Variant
    
    'get data body
    If ExcludeFilteredRows Then
      numRows = Me.VisibleRowCount
      dataArray = pSheet.Range(pSheet.Cells(DataStartCell.row, DataStartCell.column), pSheet.Cells(numRows + DataStartCell.row - 1, DataStartCell.column + TableColumnCount - 1))
      
      Set rng = pTable.DataBodyRange.SpecialCells(xlCellTypeVisible)
      'go through each area and get each row in each area
      For i = 1 To rng.Areas.count
        area = rng.Areas(i)
        'Set firstRow = rng.Areas(i).range(rng.Cells(1, 1), rng.Cells(1, rng.Columns.count))
        For row = 1 To UBound(area, 1)
          For col = 1 To UBound(area, 2)
            dataArray(currentIndex, col) = area(row, col)
          Next col
          currentIndex = currentIndex + 1
        Next row
      Next i
      
    Else
      dataArray = pTable.DataBodyRange
    End If
    

    Set FetchAllAsObjectDictionary = FetchVariantObjectDictionary(dataArray, pKey, ignoreDuplicates)

End Function


Private Function FetchVariantObjectDictionary(dataArray As Variant, Optional pKey = "", Optional ignoreDuplicates As Boolean = False) As Scripting.Dictionary
    Dim arr As Variant
    Dim field As String
    Dim obj As Scripting.Dictionary
    Dim j As Long
    Dim i As Long
    Dim tempData As Variant
    Dim rng As Range
    Dim headerRange As Variant
    
    Dim headerArray As Variant
    
    Dim primaryKey As String
    If pKey = "" Then
      primaryKey = pPrimaryKey
    Else
      primaryKey = pKey
    End If
    
    headerArray = pTable.HeaderRowRange
    Dim dataLine As Scripting.Dictionary
    Dim data As New Scripting.Dictionary
    'Set data = New Collection
    
    If Not IsEmpty(dataArray) Then
    
        For i = 1 To UBound(dataArray, 1)
            Set dataLine = New Scripting.Dictionary
            For j = LBound(dataArray, 2) To UBound(dataArray, 2)
              dataLine.Add headerArray(1, j), dataArray(i, j)
            Next j
            
            If ignoreDuplicates Then
                If Not data.Exists(CStr(dataLine(primaryKey))) Then
                    data.Add CStr(dataLine(primaryKey)), dataLine
                End If
            Else
                data.Add CStr(dataLine(primaryKey)), dataLine
            End If
        Next i
    End If
    Set FetchVariantObjectDictionary = data

End Function


'----------------------------------------------------------------------------------------------------
'*
'* FetchSelectedRowsAsObjectDictionary
'*
'----------------------------------------------------------------------------------------------------
Function FetchSelectedRowsAsObjectDictionary(Optional pKey = "", Optional ignoreDuplicates As Boolean = False)
  Dim dataArray As Variant
  
  Dim vKey As Variant
  
  Dim indexes As Scripting.Dictionary
  Set indexes = SelectedIndexes
  Dim dict As Scripting.Dictionary

  Dim vData As Variant
  Dim currentIndex As Long
  Dim firstRow As Range
  Dim index As Long
  Dim col As Long
  currentIndex = 1
  vData = pTable.DataBodyRange
  Dim columnCount As Long
  columnCount = UBound(vData, 2)
  
  If indexes.count > 0 Then
    
    dataArray = pSheet.Range(pSheet.Cells(DataStartCell.row, DataStartCell.column), pSheet.Cells(indexes.count + DataStartCell.row - 1, DataStartCell.column + TableColumnCount - 1))
    Set firstRow = pTable.DataBodyRange.Range(pTable.DataBodyRange.Cells(1, 1), pTable.DataBodyRange.Cells(1, pTable.DataBodyRange.Columns.count))
    
    For Each vKey In indexes
      index = CInt(vKey)
      For col = 1 To columnCount
        dataArray(currentIndex, col) = vData(index, col)
      Next col
      currentIndex = currentIndex + 1
    
    Next vKey
  End If 'end if numRows > 0
  
  
  
  Set FetchSelectedRowsAsObjectDictionary = FetchVariantObjectDictionary(dataArray, pKey, ignoreDuplicates)
End Function


'----------------------------------------------------------------------------------------------------
'*
'* FetchSelectedRowAsObjectDictionary
'*
'----------------------------------------------------------------------------------------------------
Function FetchSelectedRowAsObjectDictionary()
  Dim obj As Scripting.Dictionary: Set obj = Nothing
  Dim tableRow As Long
  Dim indexes As Scripting.Dictionary
  Set indexes = SelectedIndexes
  Dim var As Variant
  
  If indexes.count = 1 Then
    For Each var In indexes
        tableRow = CInt(var)
        Set obj = GetObjFromTableRow(tableRow)
    
    Next var
  End If
  
  Set FetchSelectedRowAsObjectDictionary = obj
End Function

'----------------------------------------------------------------------------------------------------
'*
'* SelectedIndexes
'*
'----------------------------------------------------------------------------------------------------
Function SelectedIndexes() As Scripting.Dictionary
  
  Dim tableRange As Range
  Set tableRange = pTable.DataBodyRange
  
  Dim selectedRange As Range
  Set selectedRange = Selection
  
  Dim unionRange As Range
  Set unionRange = Intersect(tableRange, selectedRange)
  
  Dim rng As Range
  Dim indexes As Scripting.Dictionary
  Set indexes = New Scripting.Dictionary
    
  Dim index As Long
  Dim startRow As Long
  
  If Not unionRange Is Nothing Then
    startRow = Me.DataStartCell.row
    For Each rng In unionRange.Cells
        index = rng.row - startRow + 1
        If Not indexes.Exists(index) Then
          indexes.Add index, True
        End If
    Next rng
  End If
  
  Set SelectedIndexes = indexes
End Function




'----------------------------------------------------------------------------------------------------
'*
'* FetchAsObjectCollection
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllAsObjectCollection() As Collection
    Dim arr As Variant
    Dim field As String
    Dim obj As Scripting.Dictionary
    Dim j As Long
    Dim i As Long
    
    Dim headerArray As Variant
    Dim dataArray As Variant
    
    'get header array
    headerArray = pTable.HeaderRowRange
  
    'get data body
    dataArray = pTable.DataBodyRange
    Dim dataLine As Scripting.Dictionary
    Dim data As New Collection
    Set data = New Collection
    
    For i = 1 To UBound(dataArray, 1)
        Set dataLine = New Scripting.Dictionary
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
          dataLine.Add headerArray(1, j), dataArray(i, j)
        Next j
        data.Add dataLine
    Next i


    Set FetchAllAsObjectCollection = data

End Function


'----------------------------------------------------------------------------------------------------
'*
'* GetObjFromSheetRow
'*
'----------------------------------------------------------------------------------------------------
Function GetObjFromSheetRow(row As Long) As Scripting.Dictionary
  Dim tableRow As Long
  
  tableRow = row - DataStartCell.row + 1
  
  Set GetObjFromSheetRow = GetObjFromTableRow(tableRow)
End Function



'----------------------------------------------------------------------------------------------------
'*
'* GetObjFromTableRow
'*
'----------------------------------------------------------------------------------------------------
Function GetObjFromTableRow(row As Long) As Scripting.Dictionary
  Dim dataLine As Scripting.Dictionary
  Dim dataArray As Variant
  Dim headerArray As Variant
  Dim i As Long
  headerArray = pTable.HeaderRowRange
  
  If row <= DataRowCount Then
    dataArray = RowVariantArray(row)
    
    Set dataLine = New Scripting.Dictionary
    For i = LBound(dataArray, 2) To UBound(dataArray, 2)
      dataLine.Add headerArray(1, i), dataArray(1, i)
    Next i

  Else
    Set dataLine = Nothing
  End If

  Set GetObjFromTableRow = dataLine
End Function


'----------------------------------------------------------------------------------------------------
'*
'* SaveObjectDictionary
'*
'----------------------------------------------------------------------------------------------------
Function SaveObjectDictionary(dict As Scripting.Dictionary, Optional keepFilters As Boolean = False, Optional bResizeTable As Boolean = True)
    Dim vKey As Variant
    Dim col As Collection
    Set col = New Collection
    
    For Each vKey In dict
        col.Add dict(vKey)
    Next vKey
    
    SaveObjectCollection col, keepFilters, bResizeTable
End Function


Function VisibleRowCount() As Long
  VisibleRowCount = KeywordsTable.Table.ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).Cells.count
End Function

Function Table() As ListObject
  Set Table = pTable
End Function

'----------------------------------------------------------------------------------------------------
'*
'* SaveObjectCollection
'*
'----------------------------------------------------------------------------------------------------
Function SaveObjectCollection(col As Collection, Optional keepFilters As Boolean = False, Optional bResizeTable As Boolean = True)
  Dim headers As Variant
  Dim headerDict As Scripting.Dictionary
  Dim i As Long
  Dim obj As Scripting.Dictionary
  Dim field As Variant
  Dim row As Long
  Set headerDict = New Scripting.Dictionary
  
  '
  ' clear table
  '
  ClearTable
  
  
  
  If col.count = 0 Then
    If bResizeTable Then
        resizeTable
    End If
    Exit Function
  End If

  headers = HeaderVariantArray(True)
  
  For i = LBound(headers, 2) To UBound(headers, 2)
    headerDict.Add headers(1, i), i
  Next i

  Dim outputArray As Variant
  If pTable.DataBodyRange Is Nothing Then
    outputArray = pTable.HeaderRowRange.Cells(1, 1).Offset(1, 0).Resize(col.count, headerDict.count)
  Else
    outputArray = pTable.DataBodyRange.Cells(1, 1).Resize(col.count, headerDict.count)
  End If

  Dim vKey As Variant
  row = 1
  For Each vKey In col
    Set obj = vKey
    For Each field In headerDict
      outputArray(row, headerDict(field)) = obj(CStr(field))
    Next field
    row = row + 1
  Next vKey

  If pTable.DataBodyRange Is Nothing Then
    pTable.HeaderRowRange.Cells(1, 1).Offset(1, 0).Resize(col.count, headerDict.count) = outputArray
  Else
    pTable.DataBodyRange.Cells(1, 1).Resize(col.count, headerDict.count) = outputArray
  End If
 
  If bResizeTable Then
     resizeTable
  End If
   
  '
  ' save filters
  '
  If keepFilters Then
    pTable.AutoFilter.ApplyFilter
  Else
    pTable.AutoFilter.ShowAllData
  End If


End Function



'----------------------------------------------------------------------------------------------------
'*
'* SetBackgroundColor
'*
'----------------------------------------------------------------------------------------------------
Function SetBackgroundColor(color As Long, Optional fillTable As Boolean = True, Optional fillBelow As Boolean = False)
    Dim lastCell As Long
    Dim rng As Range
    Dim lastTableRow As Long
    
    If Me.Table.ListRows.count = 0 Then
        lastTableRow = Me.Table.HeaderRowRange.row + 1
    Else
        lastTableRow = Me.Table.HeaderRowRange.row + Me.Table.ListRows.count
    End If
    
    ClearBackGroundColor
    If Not Me.Table.DataBodyRange Is Nothing Then
        
        If fillTable Then
            Me.Table.DataBodyRange.Interior.color = color
        End If
        If fillBelow Then
            lastCell = Me.LastRow
            Set rng = Me.Table.HeaderRowRange.Offset(Me.Table.ListRows.count + 1).Resize(lastCell - lastTableRow)
            rng.Interior.color = color
        End If
    Else
        If fillTable Then
            Me.Table.HeaderRowRange.Offset(1).Resize(1).Interior.color = color
        End If
        
        If fillBelow Then
            lastCell = Me.LastRow
            Set rng = Me.Table.HeaderRowRange.Offset(1).Resize(lastCell - lastTableRow)
            rng.Interior.color = color
        End If
    End If
    
    
    
End Function


Function ClearBackGroundColor()
    Dim rng As Range
    Dim lastCell As Long
    
    
    lastCell = Me.LastRow
    
    Set rng = Me.Table.HeaderRowRange.Offset(1).Resize(lastCell - Me.Table.HeaderRowRange.row)
    rng.Interior.ColorIndex = 0
    
End Function

Function SetColumnBackgroundColor(columnName As String, color As Long)
    Dim columnIndex As Integer
    Dim rng As Range
    columnIndex = GetColumnIndex(columnName)
    
    If columnIndex > 0 Then
        If Me.Table.ListRows.count > 0 Then
            Set rng = Me.Table.ListColumns(columnIndex).DataBodyRange
        Else
            Set rng = Me.Table.HeaderRowRange(1, columnIndex).Offset(1, 0)
        End If
        
        rng.Interior.color = color
    End If
End Function

Function SetLockFlagOn(Optional tableCells As Boolean = True, Optional toEndOfSheet As Boolean = True)
    Dim rng As Range
    Dim lastCell As Long
    
    If toEndOfSheet Then
        lastCell = Me.LastRow
        Set rng = Me.Table.HeaderRowRange.Offset(1).Resize(lastCell - Me.Table.HeaderRowRange.row)
        rng.Locked = True
    Else
        If Me.Table.DataBodyRange Is Nothing Then
            Me.Table.HeaderRowRange.Offset(1).Resize(1).Locked = True
        Else
            Me.Table.DataBodyRange.Locked = True
        End If
    End If
End Function


Function SetLockFlagOff(Optional tableCells As Boolean = True, Optional toEndOfSheet As Boolean = True)
    Dim rng As Range
    Dim lastCell As Long
    
    If toEndOfSheet Then
        lastCell = Me.LastRow
        Set rng = Me.Table.HeaderRowRange.Offset(1).Resize(lastCell - Me.Table.HeaderRowRange.row)
        rng.Locked = False
    Else
        If Me.Table.DataBodyRange Is Nothing Then
            Me.Table.HeaderRowRange.Offset(1).Resize(1).Locked = False
        Else
            Me.Table.DataBodyRange.Locked = False
        End If
    End If
End Function

Function SetColumnLockFlagOn(columnName As String, Optional tableCells As Boolean = True, Optional toEndOfSheet As Boolean = True)
    Dim columnIndex As Integer
    Dim rng As Range
    Dim rowCount As Long
    Dim rowsToEnd As Long
    Dim LastRow As Long
    
    rowCount = Me.Table.ListRows.count
    LastRow = Me.Table.Range(Me.Table.Range.Rows.count, 1) + Me.Table.HeaderRowRange(1, 1).row
    rowsToEnd = Me.LastRow - LastRow - 1
    columnIndex = GetColumnIndex(columnName)
    

    If columnIndex > 0 Then
        If rowCount > 0 Then
            If tableCells Then
                Set rng = Me.Table.ListColumns(columnIndex).DataBodyRange
                rng.Locked = True
            End If
            
            If toEndOfSheet And rowsToEnd > 0 Then
                Set rng = Me.Table.HeaderRowRange(1, columnIndex).Offset(rowCount + 1, 0).Resize(rowsToEnd)
                rng.Locked = True
            End If
        Else
            Set rng = Me.Table.HeaderRowRange(1, columnIndex).Offset(1, 0)
        End If
        
        
    End If


'    Dim rng As Range
'    Dim lastCell As Long
'
'    If toEndOfSheet Then
'        lastCell = Me.sheet.Rows.Count
'        Set rng = Me.Table.HeaderRowRange.Offset(1).Resize(lastCell - Me.Table.HeaderRowRange.row)
'        rng.Locked = True
'    Else
'        If Me.Table.DataBodyRange Is Nothing Then
'            Me.Table.HeaderRowRange.Offset(1).Resize(1).Locked = True
'        Else
'            Me.Table.DataBodyRange.Locked = True
'        End If
'    End If
End Function

Function SetColumnLockFlagOff(columnName As String, Optional tableCells As Boolean = True, Optional toEndOfSheet As Boolean = True)
    Dim columnIndex As Integer
    Dim rng As Range
    Dim rowCount As Long
    Dim rowsToEnd As Long
    Dim LastRow As Long
    
    rowCount = Me.Table.ListRows.count
    LastRow = Me.Table.Range(Me.Table.Range.Rows.count, 1).row
    rowsToEnd = Me.LastRow - LastRow
    columnIndex = GetColumnIndex(columnName)
    
    
    If columnIndex > 0 Then
        If rowCount > 0 Then
            If tableCells Then
                Set rng = Me.Table.ListColumns(columnIndex).DataBodyRange
                rng.Locked = False
            End If
            
            If toEndOfSheet And rowsToEnd > 0 Then
                Set rng = Me.Table.HeaderRowRange(1, columnIndex).Offset(rowCount + 1, 0).Resize(rowsToEnd)
                rng.Locked = False
            End If
        Else
            Set rng = Me.Table.HeaderRowRange(1, columnIndex).Offset(1, 0)
        End If
        
        
    End If
End Function










Function ClearFilters()
  pTable.AutoFilter.ShowAllData
End Function


'Function SaveFilters() As Scripting.Dictionary
'  Dim tableFilters As filters
'  Dim f As filter
'  Dim filters As Scripting.Dictionary
'  Dim i As Integer
'  Dim dictFilter As Scripting.Dictionary
'  Set tableFilters = pTable.AutoFilter.filters
'
'  Set filters = New Scripting.Dictionary
'
'  For i = 1 To tableFilters.count
'    Set f = tableFilters(i)
'    If f.On Then
'      Set dictFilter = New Scripting.Dictionary
'      dictFilter.Add "Criteria1", f.Criteria1
'
'      'dictFilter.Add "Criteria2", f.Criteria2
'      dictFilter.Add "Operator", f.Operator
'      dictFilter.Add "Column", i
'      filters.Add i, dictFilter
'    End If
'  Next i
'
'  Set SaveFilters = filters
'End Function
'
'Function LoadFilters(filters As Scripting.Dictionary)
'  Dim vKey As Variant
'  Dim f As Scripting.Dictionary
'
'  For Each vKey In filters
'    Set f = filters(vKey)
'    pTable.range.AutoFilter field:=f("Column"), Criteria1:=f("Criteria1"), Operator:=f("Operator")
'  Next vKey
'End Function


'Ultimiate Guide to Filters
'https://www.excelcampus.com/vba/macros-filters-autofilter-method/
'
'Function GetFilters()
'  Dim tableFilters As Filters
'  Set tableFilters = pTable.AutoFilter.Filters
'
'
'End Function
'
'Function ClearFilters()
'  pTable.AutoFilter.ShowAllData
'End Function


