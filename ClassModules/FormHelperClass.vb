Option Explicit

Function GetAutoFillDictionary(frm As Object) As Scripting.Dictionary
    Dim ctrl As Control
    Dim name As String
    Dim fieldType As String
    Dim fieldName As String
    Dim index As Long
    Dim obj As Scripting.Dictionary: Set obj = New Scripting.Dictionary
    Dim carbon As New CarbonClass
    

    For Each ctrl In frm.Controls
        name = ctrl.name
        index = InStr(1, name, "_")
        If index > 0 Then
            fieldType = Left(name, index - 1)
            fieldName = Right(name, Len(name) - index)

            If Not obj.Exists(fieldName) Then
                Select Case fieldType
                    Case "txt"
                        obj(fieldName) = ctrl.value
                    Case "chk"
                        obj(fieldName) = ctrl.value
                    Case "cmb"
                        If ctrl.ListIndex > -1 Then
                            obj(fieldName) = ctrl.List(ctrl.ListIndex, 0)
                        End If
                    Case "dte"
                       obj(fieldName) = carbon.DateFromString(ctrl.value, "dd/mm/yyyy")
                    Case Else
                        'ignore
                End Select
            End If
        End If
    Next ctrl

    Set GetAutoFillDictionary = obj
End Function



Function AutoFillFromDictionary(frm As Object, obj As Scripting.Dictionary)
    Dim ctrl As Control
    Dim name As String
    Dim fieldType As String
    Dim fieldName As String
    Dim index As Long
    'Dim obj As Scripting.Dictionary: Set obj = New Scripting.Dictionary
    Dim carbon As New CarbonClass
    

    For Each ctrl In frm.Controls
        name = ctrl.name
        index = InStr(1, name, "_")
        If index > 0 Then
            fieldType = Left(name, index - 1)
            fieldName = Right(name, Len(name) - index)

            If obj.Exists(fieldName) Then
                
                Select Case fieldType
                    Case "txt"
                        ctrl.value = obj(fieldName)
                    Case "chk"
                        ctrl.value = obj(fieldName)
                    Case "cmb"
                        On Error Resume Next
                        ctrl.value = obj(fieldName)
                        On Error GoTo 0
                    Case "dte"
                       ctrl.value = format(obj(fieldName), "dd/mm/yyyy")
                    Case Else
                        'ignore
                End Select
            End If
        End If
    Next ctrl

End Function

