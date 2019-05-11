Option Explicit
Public Const DEV_MODE = "DEV"
Public Const PROD_MODE = "LIVE"

Private pSettings As SettingsClass

Private pDBHandler As DBHandlerClass

Private pContractsTable As ExcelTablePrototype
Private pUsersTable As ExcelTablePrototype
Private pKpiTable As ExcelTablePrototype
Private pCategoryTable As ExcelTablePrototype
Private pContractKpisTable As ExcelTablePrototype
Private pHeadersTable As ExcelTablePrototype
Private pRelatedIdsTable As ExcelTablePrototype
Private pShowContractKpisTable As ExcelTablePrototype
Private pResultRowsTable As ExcelTablePrototype
Private pPeriodTable As ExcelTablePrototype
Private pHeadersReportTable As ExcelTablePrototype

Private pContractsDBTable As DBTablePrototype
Private pUsersDBTable As DBTablePrototype
Private pKpisDBTable As DBTablePrototype
Private pContractKpisDBTable As DBTablePrototype
Private pHeadersDBTable As DBTablePrototype
Private pResultRowsDBTable As DBTablePrototype
Private pPeriodDBTable As DBTablePrototype

Private pArrayFuntions As ArrayFunctionsClass
Private pFieldMaps As MapClass

Private pColorTheme As ColorThemeClass

'Menus
Private pContractsMenu As MenuHelperClass
Private pPeriodsPopup As MenuHelperClass
Private pPeriodHeadersPopup As MenuHelperClass
Private pKpisPopup As MenuHelperClass
Private pContractKpisPopup As MenuHelperClass

'----------------------------------------------------------------------------------------------------
'*
'* Generic Classes
'*
'----------------------------------------------------------------------------------------------------
Function Settings() As SettingsClass
  'If pSettings Is Nothing Then
    Set pSettings = New SettingsClass
    pSettings.Initialise wksSettings, "tblSettings"
  'End If
  Set Settings = pSettings
End Function

Function ArrayFunctions() As ArrayFunctionsClass
  If pArrayFuntions Is Nothing Then
    Set pArrayFuntions = New ArrayFunctionsClass
  End If
  Set ArrayFunctions = pArrayFuntions
End Function


'----------------------------------------------------------------------------------------------------
'*
'* Excel Tables
'*
'----------------------------------------------------------------------------------------------------
Function ContractsTable() As ExcelTablePrototype
  If pContractsTable Is Nothing Then
    Set pContractsTable = NewExcelTablePrototype(wksContracts, "tblContracts", "ContractId")
    pContractsTable.LastRow = Settings.GetSetting("CONTRACTS_LAST_ROW")
  End If
  Set ContractsTable = pContractsTable
End Function

Function UsersTable() As ExcelTablePrototype
  If pUsersTable Is Nothing Then
    Set pUsersTable = NewExcelTablePrototype(wksUsers, "tblUsers", "UserId")
    pUsersTable.LastRow = Settings.GetSetting("USERS_LAST_ROW")
  End If
  Set UsersTable = pUsersTable
End Function

Function KpisTable() As ExcelTablePrototype
  If pKpiTable Is Nothing Then
    Set pKpiTable = NewExcelTablePrototype(wksKPIs, "tblKpis", "KpiId")
    pKpiTable.LastRow = Settings.GetSetting("KPIS_LAST_ROW")
  End If
  Set KpisTable = pKpiTable
End Function

Function CategoryTable() As ExcelTablePrototype
  If pCategoryTable Is Nothing Then
    Set pCategoryTable = NewExcelTablePrototype(wksCategories, "tblCategories", "Category")
  End If
  Set CategoryTable = pCategoryTable
End Function

Function ContractKpisTable() As ExcelTablePrototype
  If pContractKpisTable Is Nothing Then
    Set pContractKpisTable = NewExcelTablePrototype(wksContractKPIs, "tblContractKpis", "KPI")
    pContractKpisTable.LastRow = Settings.GetSetting("CONTRACT_KPIS_LAST_ROW")
  End If
  Set ContractKpisTable = pContractKpisTable
End Function

Function HeadersTable() As ExcelTablePrototype
  If pHeadersTable Is Nothing Then
    Set pHeadersTable = NewExcelTablePrototype(wksPeriodHeaders, "tblResultHeaders", "KPI")
    pHeadersTable.LastRow = Settings.GetSetting("PERIOD_HEADERS_LAST_ROW")
  End If
  Set HeadersTable = pHeadersTable
End Function

Function RelatedIdsTable() As ExcelTablePrototype
  If pRelatedIdsTable Is Nothing Then
    Set pRelatedIdsTable = NewExcelTablePrototype(wksRelatedIds, "tblRelatedIds", "RelatedId")
  End If
  Set RelatedIdsTable = pRelatedIdsTable
End Function

Function ShowContractKpisTable() As ExcelTablePrototype
  If pShowContractKpisTable Is Nothing Then
    Set pShowContractKpisTable = NewExcelTablePrototype(wksShowContractKpis, "tblShowContractKpis", "KPI")
  End If
  Set ShowContractKpisTable = pShowContractKpisTable
End Function

Function ResultRowsTable() As ExcelTablePrototype
  If pResultRowsTable Is Nothing Then
    Set pResultRowsTable = NewExcelTablePrototype(wksKPIResultRows, "tblResultRows", "")
  End If
  Set ResultRowsTable = pResultRowsTable
End Function

Function PeriodTable() As ExcelTablePrototype
  If pPeriodTable Is Nothing Then
    Set pPeriodTable = NewExcelTablePrototype(wksPeriods, "tblPeriods", "Period")
    pPeriodTable.LastRow = Settings.GetSetting("PERIODS_LAST_ROW")
  End If
  Set PeriodTable = pPeriodTable
End Function

Function HeadersReportTable() As ExcelTablePrototype
  If pHeadersReportTable Is Nothing Then
    Set pHeadersReportTable = NewExcelTablePrototype(wksHeadersReport, "tblHeadersReport", "HeaderId")
  End If
  Set HeadersReportTable = pHeadersReportTable
End Function



'----------------------------------------------------------------------------------------------------
'*
'* Database Tables
'*
'----------------------------------------------------------------------------------------------------
Function ContractsDBTable() As DBTablePrototype
  If pContractsDBTable Is Nothing Then
    Set pContractsDBTable = getDbTable("Contracts", "ContractId", vbString)
  End If
  Set ContractsDBTable = pContractsDBTable
End Function


Function UsersDBTable() As DBTablePrototype
  If pUsersDBTable Is Nothing Then
    Set pUsersDBTable = getDbTable("Users", "UserId")
  End If
  Set UsersDBTable = pUsersDBTable
End Function

Function KpiDBTable() As DBTablePrototype
  If pKpisDBTable Is Nothing Then
    Set pKpisDBTable = getDbTable("Kpis", "KpiId")
  End If
  Set KpiDBTable = pKpisDBTable
End Function

Function ContractKpisDBTable() As DBTablePrototype
  If pContractKpisDBTable Is Nothing Then
    Set pContractKpisDBTable = getDbTable("ContractKpis", "ContractKpiId")
  End If
  Set ContractKpisDBTable = pContractKpisDBTable
End Function

Function HeadersDBTable() As DBTablePrototype
  If pHeadersDBTable Is Nothing Then
    Set pHeadersDBTable = getDbTable("ResultHeaders", "HeaderId")
  End If
  Set HeadersDBTable = pHeadersDBTable
End Function

Function ResultRowsDBTable() As DBTablePrototype
  If pResultRowsDBTable Is Nothing Then
    Set pResultRowsDBTable = getDbTable("ResultRows", "ResultId")
  End If
  Set ResultRowsDBTable = pResultRowsDBTable
End Function

Function PeriodDBTable() As DBTablePrototype
  If pPeriodDBTable Is Nothing Then
    Set pPeriodDBTable = getDbTable("Periods", "PeriodId")
  End If
  Set PeriodDBTable = pPeriodDBTable
End Function


'----------------------------------------------------------------------------------------------------
'*
'* Field Mapping
'*
'----------------------------------------------------------------------------------------------------
Function FieldMapping() As MapClass
    If pFieldMaps Is Nothing Then
        Set pFieldMaps = New MapClass
    End If
    Set FieldMapping = pFieldMaps
End Function


'----------------------------------------------------------------------------------------------------
'*
'* Menus
'*
'----------------------------------------------------------------------------------------------------
Function ContractsMenu() As MenuHelperClass
    If pContractsMenu Is Nothing Then
        Set pContractsMenu = NewContractsMenu()
    End If
    Set ContractsMenu = pContractsMenu
End Function

Function PeriodsPopup() As MenuHelperClass
    If pPeriodsPopup Is Nothing Then
        Set pPeriodsPopup = NewPeriodsPopup()
    End If
    Set PeriodsPopup = pPeriodsPopup
End Function

Function PeriodHeadersPopup() As MenuHelperClass
    If pPeriodHeadersPopup Is Nothing Then
        Set pPeriodHeadersPopup = NewPeriodHeadersPopup()
    End If
    Set PeriodHeadersPopup = pPeriodHeadersPopup
End Function

Function KpisPopup() As MenuHelperClass
    If pKpisPopup Is Nothing Then
        Set pKpisPopup = NewKpisPopup()
    End If
    Set KpisPopup = pKpisPopup
End Function

Function ContractKpisPopup() As MenuHelperClass
    If pContractKpisPopup Is Nothing Then
        Set pContractKpisPopup = NewContractKpisPopup()
    End If
    Set ContractKpisPopup = pContractKpisPopup
End Function


'----------------------------------------------------------------------------------------------------
'*
'* Color Theme
'*
'----------------------------------------------------------------------------------------------------
Function ColorTheme() As ColorThemeClass
    If pColorTheme Is Nothing Then
        Set pColorTheme = New ColorThemeClass
    End If
    Set ColorTheme = pColorTheme
End Function


'----------------------------------------------------------------------------------------------------
'*
'* Functions
'*
'----------------------------------------------------------------------------------------------------
Function DBHandler() As DBHandlerClass
  If pDBHandler Is Nothing Then
    Set pDBHandler = New DBHandlerClass
    pDBHandler.InitialiseConnection
  End If
  Set DBHandler = pDBHandler
End Function

Function getDbTable(name As String, Optional primaryKey As String = "", Optional primaryKeyType As VbVarType = VbVarType.vbLong) As DBTablePrototype
  Dim Table As New DBTablePrototype
  Set Table = New DBTablePrototype
  Table.InitialiseTable DBHandler(), name, primaryKey, primaryKeyType
  
  Set getDbTable = Table
End Function

Function NewExcelTablePrototype(wks As Worksheet, tableName As String, Optional primaryKey As String = "") As ExcelTablePrototype
  Dim obj As ExcelTablePrototype
  Set obj = New ExcelTablePrototype
  obj.InitialiseSheet wks, tableName, primaryKey
  Set NewExcelTablePrototype = obj
End Function




