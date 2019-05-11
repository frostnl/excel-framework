Option Explicit

Private pMenuBarPosition As MsoBarPosition
Private pMenuName As String
Private pBar As Office.CommandBar

' This type structure holds the data for a single command bar. The elements
' are listed in the order in which they appear in the wksCommandBars table.
Private Type COMMANDBAR_PROPERTIES
    sBarName As String      ' The name of the CommandBar.
    lPosition As Long       ' The location of the CommandBar.
    bIsMenuBar As Boolean   ' Whether or not the CommandBar will be a menu bar.
    bVisible As Boolean     ' Whether or not the CommandBar will be made immediately visible.
    lWidth As Long          ' You can specify a width for msoBarFloating command bars.
    lProtection As Long     ' Controls what kinds of changes the user will be allowed to make to the CommandBar.
    bIsTemporary As Boolean ' Whether the CommandBar will persist between sessions.
    bIsEnabled As Boolean   ' Whether the CommandBar will be enabled upon creation. Disabled CommandBars are not visible to the user.
End Type

' This type structure holds the data for a single command bar control.
' The elements are listed in the order in which they appear in the wksCommandBars table.
Private Type CONTROL_PROPERTIES
    sControlName As String  ' The name of the control.
    lWidth As Long          ' The width of the control.
    bIsTemporary As Boolean ' Whether the control will persist between sessions.
    bIsEnabled As Boolean   ' Whether the control will be enabled upon creation.
    sOnAction As String     ' The macro assigned to the control.
    lControlID As Long      ' Used to specify a built-in control.
    lControlType As Long    ' What kind of control this is.
    lControlStyle As Long   ' Applies only to controls of lControlType msoControlButton. Specifies the appearance of the control.
    vFaceID As Variant      ' Used to specify the control face to be used.
    bBeginGroup As Boolean  ' Whether this control has a separator bar above/left of it.
    lBefore As Long         ' The index of the control to add the control before.
    sTooltip As String      ' The tootip for this control.
    sShortcutKey As String  ' The shortcut key, if any. This just displays the shortcut key. The shortcut key must be *set* in the caption.
    sTag As String          ' String data type storage for the programmer's use.
    vParameter As Variant   ' Variant data type storage for the programmer's use.
    lState As Long          ' Specifies whether the button should be depressed or normal upon creation.
    rngListRange As Excel.Range   ' The list used to populate dropdown and combobox controls.
End Type



Function NewMenu(menuPosition As MsoBarPosition, menuName As String)
    'Dim bValidMenu As Boolean
    'bValidMenu = False

    If BarExists(menuName) Then
        Application.CommandBars(menuName).Delete
    End If

    Select Case menuPosition
        Case msoBarPopup
            pMenuBarPosition = menuPosition
            pMenuName = menuName
            Set pBar = AddPopupBar(pMenuName)
        Case Else
        
    End Select

    'If bValidMenu Then

    
        
    'End If

End Function


Function Show()
    Dim cbrBar As Office.CommandBar
    
    ' Only attempt to display the custom right-click
    ' command bar if it exists.
'    On Error Resume Next
'        Set cbrBar = Nothing
        Set cbrBar = Application.CommandBars(pMenuName)
'    On Error GoTo 0
    
    If Not cbrBar Is Nothing Then
        ' Show our custom right-click command bar.
        cbrBar.ShowPopup
        ' Cancel the default action of the right-click.
        
    End If
End Function


Function AddControl(controlType As MsoControlType, controlStyle As MsoButtonStyle, controlName As String, onAction As String, isTemporary As Boolean)

    Dim ctlControl As Office.CommandBarControl
    
    Set ctlControl = pBar.Controls.Add(controlType, , , , isTemporary)
    
    ctlControl.Caption = controlName
    ctlControl.Style = controlStyle
    ctlControl.Enabled = True
    ctlControl.Width = 150
    If Len(onAction) > 0 Then ctlControl.onAction = onAction

    ' Return an object reference to the new control.
    'Set ctlAddNewControl = ctlControl

End Function




Private Function bCommandbarExists(ByVal sBarName As String, ByRef cbrBar As Office.CommandBar) As Boolean

    If IsNumeric(sBarName) Then
        ' If an index was passed for the CommandBar name, check for it directly.
        On Error Resume Next
            Set cbrBar = Application.CommandBars(CLng(sBarName))
        On Error GoTo 0
    Else
        ' Otherwise loop the CommandBars collection and look for a name match.
        For Each cbrBar In Application.CommandBars
            ' If a match is located, exit the loop.
            If StrComp(cbrBar.name, sBarName, vbTextCompare) = 0 Then Exit For
        Next cbrBar
    End If
    
    bCommandbarExists = Not cbrBar Is Nothing

End Function


Private Function BarExists(sBarName As String) As Boolean
    Dim cbrBar As Office.CommandBar
    Dim bExists As Boolean
    bExists = False
    
    ' Otherwise loop the CommandBars collection and look for a name match.
    For Each cbrBar In Application.CommandBars
        ' If a match is located, exit the loop.
        If StrComp(cbrBar.name, sBarName, vbTextCompare) = 0 Then
            bExists = True
            Exit For
        End If
    Next cbrBar

    BarExists = bExists
End Function

Private Function AddPopupBar(sBarName As String) As Office.CommandBar
    Dim uCommandBarAtr As COMMANDBAR_PROPERTIES
    Dim cbrBar As Office.CommandBar

        With uCommandBarAtr
            'set command bar position
            .lPosition = msoBarPopup
            .bIsMenuBar = False
            .bVisible = False
            .bIsTemporary = True
            .sBarName = sBarName
        End With
        
    uCommandBarAtr.bIsTemporary = False
        
    With uCommandBarAtr
        Set cbrBar = Application.CommandBars.Add(.sBarName, .lPosition, .bIsMenuBar, .bIsTemporary)
    End With
        
    Set AddPopupBar = cbrBar
End Function

Private Function ctlAddNewControl(ByRef objTarget As Office.CommandBar, ByRef uCtlProperties As CONTROL_PROPERTIES) As Office.CommandBarControl

End Function


