Option Explicit
Private pSettings As Scripting.Dictionary
Private pSheet As Worksheet
Private pTable As ListObject

Function Initialise(sheet As Worksheet, tableName As String)
    Set pSettings = New Scripting.Dictionary
    Set pSheet = sheet
    Set pTable = sheet.ListObjects(tableName)
    LoadSettings
End Function

Function LoadSettings()
    Dim vData As Variant
    Dim i As Long
    
    If Not pTable.DataBodyRange Is Nothing Then
        vData = pTable.DataBodyRange
        For i = LBound(vData, 1) To UBound(vData, 1)
          pSettings.Add vData(i, 1), vData(i, 2)
        Next i
    End If
    
End Function

Function GetSetting(settingName As String) As Variant
  If pSettings.Exists(settingName) Then
    GetSetting = pSettings(settingName)
  Else
    GetSetting = ""
  End If
End Function

Function SetSetting(settingName As String, val As Variant) As Variant
  Dim dict As Scripting.Dictionary
  
  Dim rng As Range
  Set rng = pTable.ListColumns(1).DataBodyRange
  
  Dim c As Range
  Set c = rng.Find(What:=settingName, LookIn:=xlValues, LookAt:=xlWhole)
  
  If Not c Is Nothing Then
    c.Offset(0, 1).value = val
  End If
  
End Function



