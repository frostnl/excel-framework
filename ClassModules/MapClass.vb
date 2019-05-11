Option Explicit

Private mapTable As ExcelTablePrototype
Private mapCollection As Collection


Function GetMap(mapName As String, Optional NameAsKey As Boolean = True)
    Dim var As Variant
    Dim row As Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    
    For Each var In mapCollection
        Set row = var
        If row("MapName") = mapName Then
            map.Add row("FieldName"), row("DatabaseField")
        End If
        
    Next var
    
    Set GetMap = map

End Function

Private Sub Class_Initialize()
    Set mapTable = NewExcelTablePrototype(wksFieldMapping, "tblFieldMapping")
    Set mapCollection = mapTable.FetchAllAsObjectCollection
    
End Sub
