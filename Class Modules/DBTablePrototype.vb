Option Explicit
Private DBHandler As DBHandlerClass
Private pTableName As String
Private pPrimaryKey As String
Private pPrimaryKeyVarType As VbVarType

'----------------------------------------------------------------------------------------------------
'*
'* InitialiseTable
'*
'----------------------------------------------------------------------------------------------------
Function InitialiseTable(handler As DBHandlerClass, tableName As String, Optional primaryKey As String = "", Optional primaryKeyVarType As VbVarType = VbVarType.vbLong)
    Set DBHandler = handler
    pTableName = tableName
    pPrimaryKey = primaryKey
    pPrimaryKeyVarType = primaryKeyVarType
End Function



'----------------------------------------------------------------------------------------------------
'*
'* UpsertAll
'*
'----------------------------------------------------------------------------------------------------
Function UpdateAll(newItems As Scripting.Dictionary, Optional primaryKey As String = "")
  Dim rs As ADODB.Recordset
  Dim vKey As Variant
  Dim item As Scripting.Dictionary
  Dim sKey As String
  Set rs = FetchAllAsRecordset
  
  If primaryKey = "" Then
    primaryKey = pPrimaryKey
  End If
  
  Do While Not rs.EOF
  sKey = CStr(rs.fields(primaryKey))
    If newItems.Exists(sKey) Then
      Set item = newItems(sKey)
      For Each vKey In item
        If CStr(vKey) <> primaryKey Then
        
          'check if field exists
          On Error Resume Next
          rs.fields(CStr(vKey)) = item(vKey)
          
        End If
        
      Next vKey
      rs.Update
      
    End If
   
    rs.MoveNext
  Loop
  CloseRecordset rs
  
  
  
End Function


'----------------------------------------------------------------------------------------------------
'*
'* Upsert To Query
'*
'----------------------------------------------------------------------------------------------------
Function UpsertAllToQuery(items As Scripting.Dictionary, sql As String, key As String, Optional bOverwrite As Boolean = True, Optional updatePrimaryKey As Boolean = False, Optional primaryKeyVarType As VbVarType = VbVarType.vbLong)
  Dim rs As ADODB.Recordset
  Dim vKey As Variant
  Dim var As Variant
  Dim item As Scripting.Dictionary
  Dim sKey As String
  Set rs = FetchSqlAsRecordset(sql)
  Dim sType As String
  Dim vPrimaryKey As Variant
  
  If primaryKeyVarType = vbLong Or primaryKeyVarType = vbInteger Then
     sType = "Numeric"
  ElseIf primaryKeyVarType = vbString Then
     sType = "String"
  End If
  
  
  Do While Not rs.EOF
    'if can be updated
    Set item = Nothing
    vPrimaryKey = ""
    If (sType = "Numeric") Then
        If items.Exists(CLng(rs.fields(key))) Then
            vPrimaryKey = CLng(rs.fields(key))
        End If
    ElseIf (items.Exists(CStr(rs.fields(key))) And sType = "String") Then
        'Set item = items(CStr(rs.fields(key)))
        vPrimaryKey = CStr(rs.fields(key))
    End If
    
    On Error Resume Next
    If vPrimaryKey <> "" Then
        Set item = items(vPrimaryKey)
        'update all
        For Each vKey In item
          If CStr(vKey) <> pPrimaryKey Then
          
            rs.fields(CStr(vKey)) = item(vKey)
            
            If Err.Number = 0 Then
                'all okay
            ElseIf Err.Number = 3265 Then
                'not found in the collection, do not force error
            Else
                On Error GoTo 0
                Err.Raise Err.Number
            End If
            Err.Number = 0
          
          
          
            
          End If
        Next vKey
        rs.Update
        items.Remove vPrimaryKey
        'items.Remove rs.fields(key)
    Else
        'remove record if it doesn't exist in the items set and if overwrite is set to true
        If bOverwrite Then
            rs.Delete
        End If
    End If
    rs.MoveNext
  Loop
  
  On Error Resume Next
  
  ' Add any new items
  For Each var In items
    Set item = items(var)
    rs.AddNew
    For Each vKey In item
      If CStr(vKey) <> pPrimaryKey Or updatePrimaryKey = True Then
        rs.fields(CStr(vKey)) = item(vKey)
        
        If Err.Number = 0 Then
            'all okay
        ElseIf Err.Number = 3265 Then
            'not found in the collection, do not force error
        Else
            On Error GoTo 0
            Err.Raise Err.Number
        End If
        Err.Number = 0
      End If
    Next vKey
    rs.Update
  Next var
  
  
  CloseRecordset rs
  
  
  
End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchAllAsRecordset
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllAsRecordset() As ADODB.Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "SELECT * FROM " & pTableName
    DBHandler.mcnConnect.Open
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open sql, DBHandler.mcnConnect
   
   Set FetchAllAsRecordset = rs
End Function

Function FetchSqlAsRecordset(sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    DBHandler.mcnConnect.Open
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open sql, DBHandler.mcnConnect
   
   Set FetchSqlAsRecordset = rs
End Function



'----------------------------------------------------------------------------------------------------
'*
'* CloseRecordset
'*
'----------------------------------------------------------------------------------------------------
Function CloseRecordset(ByRef rs As ADODB.Recordset)
  DBHandler.CleanupRecordset rs
End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchAllAsClassCollection
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllAsClassCollection(className As String) As Collection
'https://www.codeproject.com/Articles/164036/Reflection-in-VBA-a-CreateObject-function-for-VBA
    Dim col As Collection: Set col = New Collection
    Dim obj As Object
    Dim vKey As Variant
    Dim map As Scripting.Dictionary
    Dim item As String
    Dim dbFieldName As String
    Dim propertyName As String
    
    Dim rs As ADODB.Recordset: Set rs = FetchAllAsRecordset
  
    'Use each recordset item to populate dictionary
    Do While Not rs.EOF
      Set obj = CreateVBAClass(className)
      Set map = obj.fieldMap
      
      For Each vKey In map
        dbFieldName = vKey
        propertyName = map(vKey)
        If Not IsNull(rs.fields(dbFieldName)) Then
            CallByName obj, propertyName, VbLet, rs.fields(dbFieldName)
       End If
      Next vKey
        
      col.Add obj
      rs.MoveNext
    Loop

    CloseRecordset rs
  Set FetchAllAsClassCollection = col
End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchAllAsObjectCollection
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllAsObjectCollection() As Collection
'https://www.codeproject.com/Articles/164036/Reflection-in-VBA-a-CreateObject-function-for-VBA
    Dim col As Collection: Set col = New Collection
    Dim obj As Scripting.Dictionary
    Dim vKey As Variant
    Dim map As Scripting.Dictionary
    Dim item As String
    Dim dbFieldName As String
    Dim propertyName As String
    Dim field As ADODB.field
    Dim rs As ADODB.Recordset: Set rs = FetchAllAsRecordset
  
    'Use each recordset item to populate dictionary
    Do While Not rs.EOF
      Set obj = New Scripting.Dictionary
      For Each field In rs.fields
        If Not IsNull(rs.fields(field.name)) Then
           obj.Add field.name, rs.fields(field.name).value
        Else
           obj.Add field.name, ""
        End If
      Next field
 
      col.Add obj
      rs.MoveNext
    Loop
  DBHandler.CleanupRecordset rs

  Set FetchAllAsObjectCollection = col
End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchAllAsObjectDictionary
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllAsObjectDictionary(Optional primaryKey = "") As Scripting.Dictionary
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    Dim obj As Scripting.Dictionary

    Dim field As ADODB.field
    Dim rs As ADODB.Recordset: Set rs = FetchAllAsRecordset
  
    If primaryKey = "" Then
      primaryKey = pPrimaryKey
    End If
  
    'Use each recordset item to populate dictionary
    Do While Not rs.EOF
      Set obj = New Scripting.Dictionary
      For Each field In rs.fields
        If Not IsNull(rs.fields(field.name)) Then
           obj.Add field.name, rs.fields(field.name).value
        Else
           obj.Add field.name, ""
        End If
      Next field

        
      dict.Add obj(primaryKey), obj
      rs.MoveNext
    Loop
  DBHandler.CleanupRecordset rs

  Set FetchAllAsObjectDictionary = dict
End Function



'----------------------------------------------------------------------------------------------------
'*
'* AppendAllFromObjectCollection
'*
'----------------------------------------------------------------------------------------------------
Function AppendAllFromObjectCollection(col As Collection)
'https://www.codeproject.com/Articles/164036/Reflection-in-VBA-a-CreateObject-function-for-VBA

    Dim vKey As Variant
    Dim dataRow As Scripting.Dictionary
    Dim dbFieldName As String
    Dim objField As Variant
    Dim fieldName As String
    
    Dim rs As ADODB.Recordset
    Dim sql As String
      
    
    sql = "SELECT TOP 1 * FROM " & pTableName
    Set rs = DBHandler.GetRecordsetForUpdate(sql)
    
    On Error Resume Next
    
     For Each vKey In col
     Set dataRow = vKey
     rs.AddNew
       For Each objField In dataRow
         fieldName = CStr(objField)
         If fieldName <> pPrimaryKey Then
            rs.fields(fieldName) = dataRow(objField)
            
            If Err.Number = 0 Then
                'all okay
            ElseIf Err.Number = 3265 Then
                'not found in the collection, do not force error
            Else
                On Error GoTo 0
                Err.Raise Err.Number
            End If
            Err.Number = 0
         
         End If
       Next objField
     rs.Update
     Next vKey

DBHandler.CleanupRecordset rs

End Function


'----------------------------------------------------------------------------------------------------
'*
'* TruncateTable
'*
'----------------------------------------------------------------------------------------------------
Function TruncateTable()

End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchByKey
'*
'----------------------------------------------------------------------------------------------------
Function FetchByKey(value As Variant, Optional field As String = "", Optional fieldType As VbVarType)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim f As ADODB.field
    Dim obj As Scripting.Dictionary
    
    If field = "" Then
        field = pPrimaryKey
        fieldType = pPrimaryKeyVarType
    End If

    If fieldType = vbString Then
        sql = "SELECT * FROM " & pTableName & " WHERE " & field & " = '" & value & "'"
    ElseIf fieldType = vbInteger Or fieldType = vbLong Or fieldType = vbDouble Then
        sql = "SELECT * FROM " & pTableName & " WHERE " & field & " = " & value
    ElseIf fieldType = vbDate Then
        sql = "SELECT * FROM " & pTableName & " WHERE " & field & " = #" & year(value) & "-" & month(value) & "-" & day(value) & "#"
    
    End If
    
    Set rs = FetchSqlAsRecordset(sql)
    
     If Not rs.EOF Then
      Set obj = New Scripting.Dictionary
      For Each f In rs.fields
        If Not IsNull(rs.fields(f.name)) Then
           obj.Add f.name, rs.fields(f.name).value
        Else
           obj.Add f.name, ""
        End If
      Next f

    End If
     
    CloseRecordset rs

    Set FetchByKey = obj
End Function



'----------------------------------------------------------------------------------------------------
'*
'* FetchAllByQuery
'*
'----------------------------------------------------------------------------------------------------
Function FetchAllByQuery(sql As String)
'https://www.codeproject.com/Articles/164036/Reflection-in-VBA-a-CreateObject-function-for-VBA
    Dim col As Collection: Set col = New Collection
    Dim obj As Scripting.Dictionary
    Dim vKey As Variant
    Dim map As Scripting.Dictionary
    Dim item As String
    Dim dbFieldName As String
    Dim propertyName As String
    Dim field As ADODB.field
    Dim rs As ADODB.Recordset: Set rs = FetchSqlAsRecordset(sql)
  
    'Use each recordset item to populate dictionary
    Do While Not rs.EOF
      Set obj = New Scripting.Dictionary
      For Each field In rs.fields
        If Not IsNull(rs.fields(field.name)) Then
           obj.Add field.name, rs.fields(field.name).value
        Else
           obj.Add field.name, ""
        End If
      Next field
 
      col.Add obj
      rs.MoveNext
    Loop
  DBHandler.CleanupRecordset rs

  Set FetchAllByQuery = col
End Function



'----------------------------------------------------------------------------------------------------
'*
'* AddNew
'*
'----------------------------------------------------------------------------------------------------
Function AddNew(row As Scripting.Dictionary)
    Dim vKey As Variant
    Dim dataRow As Scripting.Dictionary
    Dim dbFieldName As String
    Dim objField As Variant
    Dim fieldName As String
    
    Dim rs As ADODB.Recordset

    Dim sql As String
    sql = "SELECT TOP 1 * FROM " & pTableName
    Set rs = DBHandler.GetRecordsetForUpdate(sql)
    
    rs.AddNew
    For Each objField In row
      fieldName = CStr(objField)
      If fieldName <> pPrimaryKey Then
        rs.fields(fieldName) = row(fieldName)
      End If
    Next objField

    rs.Update
    DBHandler.CleanupRecordset rs
End Function



'----------------------------------------------------------------------------------------------------
'*
'* UpdateExisting
'*
'----------------------------------------------------------------------------------------------------
Function UpdateExisting(data As Scripting.Dictionary)
    Dim vKey As Variant
    Dim dataRow As Scripting.Dictionary
    Dim dbFieldName As String
    Dim objField As Variant
    Dim fieldName As String
    
    Dim rs As ADODB.Recordset

    Dim sql As String
    If data.count > 0 Then
        If pPrimaryKeyVarType = vbInteger Or pPrimaryKeyVarType = vbLong Then
            sql = "SELECT TOP 1 * FROM " & pTableName & " WHERE " & pPrimaryKey & " = " & data(pPrimaryKey)
        ElseIf pPrimaryKeyVarType = vbString Then
            sql = "SELECT TOP 1 * FROM " & pTableName & " WHERE " & pPrimaryKey & " = '" & data(pPrimaryKey) & "'"
        End If
        
        Set rs = DBHandler.GetRecordsetForUpdate(sql)
        
        For Each objField In data
          fieldName = CStr(objField)
          If fieldName <> pPrimaryKey Then
            rs.fields(fieldName) = data(fieldName)
          End If
        Next objField
    
        rs.Update
        DBHandler.CleanupRecordset rs
    End If
End Function



'----------------------------------------------------------------------------------------------------
'*
'* Upsert
'*
'----------------------------------------------------------------------------------------------------
Function Upsert(data As Scripting.Dictionary)
    Dim vKey As Variant
    Dim dataRow As Scripting.Dictionary
    Dim dbFieldName As String
    Dim objField As Variant
    Dim fieldName As String
    
    Dim rs As ADODB.Recordset

    Dim sql As String
    
    If pPrimaryKeyVarType = vbInteger Or pPrimaryKeyVarType = vbLong Then
        sql = "SELECT TOP 1 * FROM " & pTableName & " WHERE " & pPrimaryKey & " = " & data(pPrimaryKey)
    ElseIf pPrimaryKeyVarType = vbString Then
        sql = "SELECT TOP 1 * FROM " & pTableName & " WHERE " & pPrimaryKey & " = '" & data(pPrimaryKey) & "'"
    End If
        
    Set rs = DBHandler.GetRecordsetForUpdate(sql)
        
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    
    On Error Resume Next
    
    For Each objField In data
      fieldName = CStr(objField)
      If fieldName <> pPrimaryKey Then
        rs.fields(fieldName) = data(fieldName)
            
        If Err.Number = 0 Then
            'all okay
        ElseIf Err.Number = 3265 Then
            'not found in the collection, do not force error
        Else
            On Error GoTo 0
            Err.Raise Err.Number
        End If
        Err.Number = 0
        
      End If
    Next objField

    rs.Update
    DBHandler.CleanupRecordset rs
    
End Function




'----------------------------------------------------------------------------------------------------
'*
'* DeleteByKey
'*
'----------------------------------------------------------------------------------------------------
Function DeleteByKey(id As Variant)
    Dim sql As String
    DBHandler.mcnConnect.Open
    sql = "DELETE * FROM " & pTableName & " WHERE " & pPrimaryKey & " = " & id
    DBHandler.mcnConnect.Execute sql
    DBHandler.mcnConnect.Close
End Function

