Option Explicit

Sub ContractKpis_Show()

    Dim contractId As String
    Dim sql As String
    Dim col As Collection
    Dim var As Variant
    Dim kpi As Scripting.Dictionary
    Dim kpis As New Scripting.Dictionary
    Dim sFormat As String
    
    contractId = wksContractKPIs.Range("ptrContractId").value
 
    sql = "SELECT C.ContractId, K.KpiName, K.Format, K.KpiId, KC.ContractKpiId, KC.BelowExpectations, KC.MetExpectations, KC.AboveExpectations FROM (Contracts As C LEFT JOIN ContractKpis As KC On C.ContractId = KC.ContractId) LEFT JOIN Kpis As K On KC.KpiId = K.KpiId WHERE C.ContractId = '" & contractId & "'"
    
    Set col = ContractsDBTable.FetchAllByQuery(sql)
        
    ' get the unique KPIs only
    For Each var In col
        Set kpi = var
        If Not kpis.Exists(kpi("KpiId")) And Not kpi("KpiId") = "" Then
            kpi("BelowExpectations") = format(kpi("BelowExpectations"), kpi("Format"))
            kpi("MetExpectations") = format(kpi("MetExpectations"), kpi("Format"))
            kpi("AboveExpectations") = format(kpi("AboveExpectations"), kpi("Format"))
            kpis.Add kpi("KpiId"), kpi
        End If
    
    Next var
    
    ContractKpisTable.SaveObjectDictionary kpis, False, False
    wksContractKPIs.Range("ptrCurrentId").value = contractId
End Sub

Sub ContractKpis_BackToContracts()
    wksContracts.Activate
End Sub


Sub ContractKpis_AddNew()
    Dim contractKpi As New Scripting.Dictionary
    Dim contractId As String
    
    contractId = wksContractKPIs.Range("ptrContractId").value
    
    contractKpi.Add "ContractKpiId", 0
    contractKpi.Add "ContractId", contractId
    ContractKpis_ShowForm contractKpi
    
End Sub

Sub ContractKpis_Edit()
    Dim kpiId As Long
    Dim kpis As Collection
    Dim contractKpi As Scripting.Dictionary
    Dim sql As String
    Dim whereClause As String
    Dim contractId As String
    Dim relatedId As String
    
    contractId = wksContractKPIs.Range("ptrContractId").value
    'relatedId = wksContractKpis.Range("ptrRelatedId").value
    Set contractKpi = ContractKpisTable.FetchSelectedRowAsObjectDictionary
    
    If Not contractKpi Is Nothing Then
        sql = "SELECT CK.*, K.KpiName FROM (ContractKpis As CK INNER JOIN Kpis As K on CK.KpiId = K.KpiId) INNER JOIN Contracts As C ON CK.ContractId = C.ContractId WHERE CK.ContractKpiId = " & contractKpi("ContractKpiId")
        
        Set kpis = ContractKpisDBTable.FetchAllByQuery(sql)
        Set contractKpi = kpis(1)
        
        ContractKpis_ShowForm contractKpi
        
    End If
    
End Sub


Sub ContractKpis_Delete()
    Dim kpiId As Long
    Dim kpis As Collection
    Dim contractKpi As Scripting.Dictionary
    Dim sql As String
    Dim whereClause As String
    Dim contractId As String
    Dim relatedId As String
    
    contractId = wksContractKPIs.Range("ptrContractId").value
    Set contractKpi = ContractKpisTable.FetchSelectedRowAsObjectDictionary
    
    If Not contractKpi Is Nothing Then
        If MsgBox("Are you sure you want to delete this KPI?", vbYesNo) = vbYes Then
            sql = "DELETE * FROM ContractKpis WHERE ContractKpiId = " & contractKpi("ContractKpiId")
            DBHandler.ExecuteQuery sql
            ContractKpis_Show
        End If
    End If
    
End Sub

Function ContractKpis_ShowForm(contractKpi As Scripting.Dictionary)
    Dim f As New FContractKpi
    
    Dim contractKpiId As Long
    Dim kpis As Collection
    Dim kpiList As Scripting.Dictionary
    Dim formHelper As New FormHelperClass
    Dim contractId As String
    
    Dim vBelowExpectations As Variant
    Dim vMetExpectations As Variant
    Dim vAboveExpectations As Variant
    Dim num As Variant
       
    contractId = contractKpi("ContractId")
    
    'Set contractKpi = ContractKpisDBTable.FetchByKey(contractKpiId, "ContractKpiId", vbLong)
    Set kpiList = KpiDBTable.FetchAllAsObjectDictionary("KpiId")
    
    formHelper.AutoFillFromDictionary f, contractKpi
    f.txtContractKpiId = contractKpi("ContractKpiId")
    f.txtKpiId = contractKpi("KpiId")
    f.txtKpiName.value = contractKpi("KpiName")
    f.txtBelowExpectations.value = contractKpi("BelowExpectations")
    f.txtMetExpectations.value = contractKpi("MetExpectations")
    f.txtAboveExpectations.value = contractKpi("AboveExpectations")
    
    f.Initialise kpiList
    f.Show

    If f.BTN_OK = True Then
        Set contractKpi = New Scripting.Dictionary
        contractKpi.Add "ContractKpiId", f.txtContractKpiId.value
        contractKpi.Add "ContractId", contractId
        contractKpi.Add "KpiId", f.txtKpiId.value
        
        'below expectations
        vBelowExpectations = format(f.txtBelowExpectations.value, "0.0")
        If IsNumeric(vBelowExpectations) Then
            contractKpi.Add "BelowExpectations", CDbl(vBelowExpectations)
        Else
            contractKpi.Add "BelowExpectations", 0
        End If
        
        'met expectations
        vMetExpectations = format(f.txtMetExpectations.value, "0.0")
        If IsNumeric(vMetExpectations) Then
            contractKpi.Add "MetExpectations", CDbl(vMetExpectations)
        Else
            contractKpi.Add "MetExpectations", 0
        End If
        
        'above expecations
        vAboveExpectations = format(f.txtAboveExpectations.value, "0.0")
        If IsNumeric(vAboveExpectations) Then
            contractKpi.Add "AboveExpectations", CDbl(vAboveExpectations)
        Else
            contractKpi.Add "AboveExpectations", 0
        End If
        
        'ContractKpis_SaveKpi contractKpi
        ContractKpisDBTable.Upsert contractKpi
    End If

    Unload f
    ContractKpis_Show
 
End Function


Sub ContractKpis_ApplyToContracts()
    Dim contracts As Collection
    Dim var As Variant
    Dim var2 As Variant
    Dim sql As String

    Dim contractId As String
    Dim defaultRow As Scripting.Dictionary
    Dim contract As Scripting.Dictionary
    Dim defaults As Collection
    Dim kpi As Scripting.Dictionary
    Dim kpis As Scripting.Dictionary
    Dim kpiDefault As New Scripting.Dictionary
    Dim kpiDefaults As New Scripting.Dictionary
    Dim user As String
    Dim permissions As String
    
    user = Settings.GetSetting("USER_ID")
    permissions = Settings.GetSetting("USER_TYPE")
    contractId = wksContractKPIs.Range("ptrCurrentId").value
    
    
    Set contracts = GetContractsToBeApplied
    
    'if valid to continue
    If contracts.Count > 0 And MsgBox("This update will apply to " & contracts.Count & " contract(s) are you sure you want to contintue", vbYesNo) = vbYes Then

        'get all default KPIs
        'Set defaults = ContractKpisTable.FetchAllAsObjectCollection
        Set defaults = ContractKpisDBTable.FetchAllByQuery("SELECT * FROM ContractKpis WHERE ContractId = '" & contractId & "'")

        'get all KPIs
        Set kpis = KpisTable.FetchAllAsObjectDictionary("KpiId", True, False)

        For Each contract In contracts
            If contract("SpecialistCommercial") = user Or permissions = "SUPER USER" Then

                Set kpiDefaults = New Scripting.Dictionary
    
                For Each defaultRow In defaults
                    Set kpiDefault = New Scripting.Dictionary
                    Set kpi = kpis(CStr(defaultRow("KpiId")))

                    kpiDefault.Add "ContractId", contract("ContractId")
                    kpiDefault.Add "KpiId", kpi("KpiId")
                    kpiDefault.Add "BelowExpectations", defaultRow("BelowExpectations")
                    kpiDefault.Add "MetExpectations", defaultRow("MetExpectations")
                    kpiDefault.Add "AboveExpectations", defaultRow("AboveExpectations")
                    kpiDefault.Add "Weight", defaultRow("Weight")
                                       
                    kpiDefaults.Add kpiDefault("KpiId"), kpiDefault
                Next defaultRow
    
                'delete existing defaults
                DBHandler.ExecuteQuery "DELETE * FROM ContractKpis WHERE ContractId = '" & contract("ContractId") & "'"
    
                sql = "SELECT * FROM ContractKpis WHERE ContractId = '" & contract("ContractId") & "'"
                ContractKpisDBTable.UpsertAllToQuery kpiDefaults, sql, "KpiId", True, True

            End If
        Next contract
        MsgBox "Done"
    End If


End Sub


Private Function GetContractsToBeApplied() As Collection
    Dim contracts As Collection
    
    Dim dateString As String
    Dim sql As String
    
    Dim contractId As String
    Dim relatedId As String
    Dim contractStatus As String
    Dim effectiveAfterDate As String
    
    'set sql where clause
    Dim whereClause As String
    whereClause = " WHERE"
    
    
    contractId = wksContractKPIs.Range("ptrApplyContractId").value
    relatedId = wksContractKPIs.Range("ptrApplyRelatedId").value
    contractStatus = wksContractKPIs.Range("ptrApplyContractStatus").value
    effectiveAfterDate = wksContractKPIs.Range("ptrApplyEffectiveAfterDate").value
    
    sql = "Select * FROM Contracts "

    'get all contracts from database, that meet the criteria
    whereClause = " WHERE "
    
    'add where clause for contractId
    If contractId <> "" Then
        sql = sql & whereClause & " ContractId = '" & contractId & "'"
        whereClause = " AND"
    End If

    'add where clause for relatedId
    If relatedId <> "" Then
        sql = sql & whereClause & " RelatedId = '" & relatedId & "'"
        whereClause = " AND"
    End If

    'add where clause for contractStatus
    Select Case contractStatus
        Case "Published", "Publishing"
            sql = sql & whereClause & " (ContractStatus = 'Published' OR ContractStatus = 'Publishing')"
            whereClause = " AND"
        Case "Draft"
            sql = sql & whereClause & " ContractStatus LIKE '%Draft%'"
            whereClause = " AND"
        Case "Published Or Draft"
            sql = sql & whereClause & " (ContractStatus = 'Published' OR ContractStatus = 'Publishing' OR ContractStatus LIKE '%Draft%')"
            whereClause = " AND"
        Case Else
            'do not add clause
    End Select

    'add clause for effective after date
    If IsDate(effectiveAfterDate) Then
        dateString = format(year(effectiveAfterDate), "0000") & "-" & format(month(effectiveAfterDate), "00") & "-" & format(day(effectiveAfterDate), "00")
        sql = sql & whereClause & " EffectiveDate >= #" & dateString & "#"
        whereClause = " AND "
    End If
    

    'get contracts
    Set contracts = ContractsDBTable.FetchAllByQuery(sql)
    
    Set GetContractsToBeApplied = contracts
End Function



'----------------------------------------------------------------------------------------------------
'*
'* Sheet Protection Functions
'*
'----------------------------------------------------------------------------------------------------
Sub ContractKpis_EditModeOn()
    ContractKpis_Unprotect
End Sub

Sub ContractKpis_EditModeOff()
    ContractKpis_Protect
End Sub

Sub ContractKpis_Protect()
    Application.ScreenUpdating = False
    wksContractKPIs.Unprotect Password:=adminPassword

        If Settings.GetSetting("ENVIRONMENT") = PROD_MODE Then
        wksContractKPIs.Protect Password:=adminPassword, DrawingObjects:=True, _
            Contents:=True, Scenarios:=False, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingColumns:=False, _
            AllowInsertingRows:=False, AllowInsertingHyperlinks:=True, _
            AllowDeletingColumns:=False, AllowDeletingRows:=False, _
            AllowSorting:=True, AllowFiltering:=True, _
            AllowUsingPivotTables:=True
    End If
    ContractKpisTable.ClearBackGroundColor
    ContractKpisTable.SetBackgroundColor ColorTheme.ThemeInactive, False, True
    Application.ScreenUpdating = True
End Sub

Sub ContractKpis_Unprotect()
    wksContractKPIs.Unprotect Password:=adminPassword
End Sub


'----------------------------------------------------------------------------------------------------
'*
'* NewContractsMenu
'*
'----------------------------------------------------------------------------------------------------
Function NewContractKpisPopup() As MenuHelperClass
    Dim menuHelper As New MenuHelperClass
    
    menuHelper.NewMenu MsoBarPosition.msoBarPopup, "ContractKpisPopup"
    menuHelper.AddControl msoControlButton, msoButtonCaption, "Edit KPI", "ContractKpis_Edit", True
    menuHelper.AddControl msoControlButton, msoButtonCaption, "Delete KPI", "ContractKpis_Delete", True
        
    Set NewContractKpisPopup = menuHelper
End Function
