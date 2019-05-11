Option Explicit

Sub Users_Show()
  Dim col As Collection
  Set col = UsersDBTable.FetchAllAsObjectCollection
  UsersTable.SaveObjectCollection col
End Sub

Sub Users_SaveSelected()
    Dim dict As Scripting.Dictionary
    Set dict = UsersTable.FetchSelectedRowsAsObjectDictionary

    If dict.Count > 0 Then
        UsersDBTable.UpdateAll dict
        MsgBox "done"
    End If
End Sub

Sub Users_UpdateUsers()
    Dim contracts As Collection
    Dim Users As Scripting.Dictionary
    Dim contract As Scripting.Dictionary
    Dim user As Scripting.Dictionary
    Dim newUsers As New Scripting.Dictionary
    Dim newUserCol As New Collection
    Dim vKey As Variant
    
    Dim cscLogon As String
    Dim cswLogon As String
    
    Dim sql As String
    sql = "SELECT * FROM CONTRACTS WHERE ContractStatus Like '%Publish%' OR ContractStatus Like '%Draft%'"
    Set contracts = ContractsDBTable.FetchAllByQuery(sql) 'only get active contracts
    
    Set Users = UsersDBTable.FetchAllAsObjectDictionary("Logon")
    
    For Each contract In contracts
        
        cscLogon = contract("SpecialistCommercialId")
        cswLogon = contract("SpecialistWorksId")
        
        If Not Users.Exists(cscLogon) And Not newUsers.Exists(cscLogon) Then
            Set user = New Scripting.Dictionary
            user.Add "Logon", cscLogon
            user.Add "UserName", contract("SpecialistCommercial")
            user.Add "UserType", "CSC"
            newUsers.Add cscLogon, user
        End If
        
        If Not Users.Exists(cswLogon) And Not newUsers.Exists(cswLogon) Then
            Set user = New Scripting.Dictionary
            user.Add "Logon", cswLogon
            user.Add "UserName", contract("SpecialistWorks")
            user.Add "UserType", "CSW"
            newUsers.Add cswLogon, user
        End If
              
    Next contract
    
    For Each vKey In newUsers
        newUserCol.Add newUsers(vKey)
    Next vKey
    
    If newUserCol.Count > 0 Then
        UsersDBTable.AppendAllFromObjectCollection newUserCol
        MsgBox newUserCol.Count & " user(s) updated"
    Else
        MsgBox "No new users added"
    End If
End Sub

Function SpecialistWorks_FetchAllAsCollection() As Collection
  Dim col As Collection
  Dim vKey As Variant
  Dim user As Scripting.Dictionary
  Dim specialists As Collection
  
  Set specialists = New Collection
  Set col = UsersTable.FetchAllAsObjectCollection()
  For Each vKey In col
    Set user = vKey
    If user("IsSpecialistWorks") = True Then
      specialists.Add user
    End If
  Next vKey
  
  Set SpecialistWorks_FetchAllAsCollection = specialists
End Function

Function SpecialistCommercial_FetchAllAsCollection() As Collection
  Dim col As Collection
  Dim vKey As Variant
  Dim user As Scripting.Dictionary
  Dim specialists As Collection
  
  Set specialists = New Collection
  Set col = UsersTable.FetchAllAsObjectCollection()
  For Each vKey In col
    Set user = vKey
    If user("IsSpecialistCommercial") = True Then
      specialists.Add user
    End If
  Next vKey
  
  Set SpecialistCommercial_FetchAllAsCollection = specialists
End Function

'----------------------------------------------------------------------------------------------------
'*
'* Sheet Protection Functions
'*
'----------------------------------------------------------------------------------------------------
Sub Users_ToggleSheetEdits()
    Dim btn As Shape
    Dim txt As String
    
    Dim status As String
    status = Settings.GetSetting("USERS_EDIT_STATUS")
    
    Set btn = wksUsers.Shapes("btnUsers_ToggleSheetEdits")

    If status = "Disabled" Then
        Users_Unlock
        'MsgBox "Enabled"
    Else
        Users_Lock
        'MsgBox "Disabled"
    End If

End Sub

Sub Users_Unlock()
    Dim btn As Shape
    Set btn = wksUsers.Shapes("btnUsers_ToggleSheetEdits")

    Users_Unprotect
    UsersTable.ClearBackGroundColor
    UsersTable.SetBackgroundColor ColorTheme.ThemeInput
    UsersTable.SetColumnBackgroundColor "UserId", ColorTheme.White
    
    UsersTable.SetLockFlagOff
    UsersTable.SetColumnLockFlagOn "UserId", True, True
    
    btn.TextFrame2.TextRange.Text = "Disable Sheet Edits"
    Settings.SetSetting "USERS_EDIT_STATUS", "Enabled"
    Users_Protect
End Sub

Sub Users_Lock()

    Dim btn As Shape
    Set btn = wksUsers.Shapes("btnUsers_ToggleSheetEdits")
    
    Users_Unprotect
    
    UsersTable.SetBackgroundColor ColorTheme.ThemeInactive, False, True
    UsersTable.SetLockFlagOn True, True
    btn.TextFrame2.TextRange.Text = "Enable Sheet Edits"
    Settings.SetSetting "Users_EDIT_STATUS", "Disabled"
    
    Users_Protect

End Sub


Sub Users_Protect()
    Application.ScreenUpdating = False
    wksUsers.Unprotect Password:=adminPassword

    If Settings.GetSetting("ENVIRONMENT") = PROD_MODE Then
        wksUsers.Protect Password:=adminPassword, DrawingObjects:=True, _
            Contents:=True, Scenarios:=False, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, AllowInsertingHyperlinks:=True, _
            AllowDeletingColumns:=False, AllowDeletingRows:=False, _
            AllowSorting:=True, AllowFiltering:=True, _
            AllowUsingPivotTables:=True
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub Users_Unprotect()
    wksUsers.Unprotect Password:=adminPassword
End Sub
