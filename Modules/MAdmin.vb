Option Explicit

Public Const adminPassword As String = "password"

Sub Admin_ProtectWorkbook()
    Periods_EditModeOff
    PeriodHeaders_EditModeOff
    Kpis_EditModeOff
    Contracts_EditModeOff
    ContractKpis_EditModeOff

End Sub

Sub Admin_UnprotectWorkbook()
    Periods_EditModeOn
    PeriodHeaders_EditModeOn
    Kpis_EditModeOn
    Contracts_EditModeOn
    ContractKpis_EditModeOff
End Sub
