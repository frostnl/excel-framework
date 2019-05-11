Option Explicit
Private rules As Collection


Function Validate(frm As FPeriod)
    


End Function



Function AddDataTypeValidation(controlName As String, varType As VbVarType)
    Dim val As New Scripting.Dictionary
    val.Add "ValidationType", "DataType"
    val.Add "ControlName", controlName
    val.Add "VarType", varType
    
    rules.Add val
End Function


Function AddFormatValidation(controlName As String, format As String)

End Function



Private Function ValidateDataType(rule As Scripting.Dictionary)
    Dim controlName As String
    Dim value As Variant
    
    controlName = rule("ControlName")
            
    


End Function



