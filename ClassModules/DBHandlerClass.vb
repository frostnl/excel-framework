Option Explicit

Public mcnConnect As New ADODB.Connection

'***************************************************************************************
'Function:      InitialiseConnection
'
'Comments:      Will initialise connection to the DQM database
'
'Date           Developer       Action
'---------------------------------------
'v1.0.0         Nigel Frost     Created
'***************************************************************************************
Function InitialiseConnection()
    
    Dim sDrive As String
    Dim sConnect As String
    Dim sFolder As String
    Dim nIMEX As Integer
    Dim sDBLocation As String
    nIMEX = 1

    
    sDrive = Left(Application.ActiveWorkbook.Path, 1)
    
    
    'Get drive as shared server may be on different drive name for different users
'    sFolder = Application.ActiveWorkbook.Path
'    If Right(sFolder, 1) <> "\" Then
'        sFolder = sFolder & "\"
'    End If

    'sDBLocation = "G:\My Drive\Gapwise Consulting\Internal Tools\Market Positioning\MarketPositioning.accdb"
    sDBLocation = Settings.GetSetting("DATABASE_LOCATION")

    sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Data Source=" & sDBLocation & ";"

    mcnConnect.CursorLocation = adUseServer
    mcnConnect.Mode = adModeReadWrite
    mcnConnect.ConnectionString = sConnect
    mcnConnect.Open
    
    'Close to allow connection pooling
    mcnConnect.Close

End Function

'***************************************************************************************
'Function:      CloseConnection
'
'Comments:      Will clear the DQM connection
'
'Date           Developer       Action
'---------------------------------------
'v1.0.0         Nigel Frost     Created
'***************************************************************************************
Function CloseConnection()
    Set mcnConnect = Nothing
End Function

'***************************************************************************************
'Function:      GetRecordsetForUpate
'
'Comments:      Will execute the SQL statement and return byref an ADODB recordset
'               with read/write access
'
'Date           Developer       Action
'---------------------------------------
'v1.0.0         Nigel Frost     Created
'***************************************************************************************
Function GetRecordsetForUpdate(sSql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    mcnConnect.Open
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open sSql, mcnConnect
    Set GetRecordsetForUpdate = rs
End Function


Function ExecuteQuery(sSql As String)
    mcnConnect.Open
    mcnConnect.Execute sSql
    mcnConnect.Close
End Function

Function RunProcedure(procedureName, ParamArray arr() As Variant)
    On Error Resume Next
    Dim appAccess As New Access.Application
    Dim loc As String
    loc = Settings.GetSetting("DATABASE_LOCATION")
    appAccess.OpenCurrentDatabase loc
    
    Select Case UBound(arr) + 1
        Case 0:
            appAccess.Run procedureName
        Case 1:
            appAccess.Run procedureName, arr(0)
        Case 2:
            appAccess.Run procedureName, arr(0), arr(1)
        Case 3:
            appAccess.Run procedureName, arr(0), arr(1), arr(2)
        Case 4:
            appAccess.Run procedureName, arr(0), arr(1), arr(2), arr(3)
        Case Else
            Err.Raise 99999
    End Select
    appAccess.CloseCurrentDatabase
    Set appAccess = Nothing
End Function

'***************************************************************************************
'Function:      CleanupRecordset
'
'Comments:      Will close and cleanup the specified ADODB recordset
'
'Date           Developer       Action
'---------------------------------------
'v1.0.0         Nigel Frost     Created
'***************************************************************************************
Function CleanupRecordset(ByRef rs As ADODB.Recordset)
    'Debug.Print "recordset state1: " & rs.State
    'Debug.Print "mcn state1: " & mcnConnect.State
    If rs.State <> adStateClosed Then rs.Close
    'Debug.Print "recordset state2: " & rs.State
    Set rs = Nothing
    Me.mcnConnect.Close
    'Debug.Print "mcn state2: " & mcnConnect.State
End Function



Function GetObjectCollectionFromSql(sql As String) As Collection
'https://www.codeproject.com/Articles/164036/Reflection-in-VBA-a-CreateObject-function-for-VBA
    Dim col As Collection: Set col = New Collection
    Dim obj As Scripting.Dictionary
    Dim vKey As Variant
    Dim map As Scripting.Dictionary
    Dim item As String
    Dim dbFieldName As String
    Dim propertyName As String
    Dim field As ADODB.field
    Dim rs As ADODB.Recordset:
    
    Set rs = GetRecordsetForUpdate(sql)
  
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
  CleanupRecordset rs

  Set GetObjectCollectionFromSql = col

End Function

Function TableExists(ByVal strTable As String) As Boolean
'an Access object
Dim objAccess As Object
'connection string to access database
Dim strConnection As String
'catalog object
Dim objCatalog As Object
'connection object
Dim cnn As Object
Dim i As Integer
Dim intRow As Integer

Dim loc As String
loc = Settings.GetSetting("DATABASE_LOCATION")

Set objAccess = CreateObject("Access.Application")
'open access database
Call objAccess.OpenCurrentDatabase(loc)
'get the connection string
strConnection = objAccess.CurrentProject.Connection.ConnectionString
'close the access project
objAccess.Quit
'create a connection object
Set cnn = CreateObject("ADODB.Connection")
'assign the connnection string to the connection object
cnn.ConnectionString = strConnection
'open the adodb connection object
cnn.Open
'create a catalog object
Set objCatalog = CreateObject("ADOX.catalog")
'connect catalog object to database
objCatalog.ActiveConnection = cnn
'loop through the tables in the catalog object
intRow = 1
For i = 0 To objCatalog.Tables.Count - 1
'check if the table is a user defined table
If objCatalog.Tables.item(i).Type = "TABLE" Then


If objCatalog.Tables.item(i).Names = strTable Then
TableExists = True
Exit Function
End If
End If
Next i

TableExists = False
End Function

