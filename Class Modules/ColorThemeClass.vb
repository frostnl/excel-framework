Option Explicit



Public Property Get Red() As Long
    Red = RGB(255, 0, 0)
End Property

Public Property Get White() As Long
    White = RGB(255, 255, 255)
End Property


Public Property Get ThemeLocked()
    ThemeLocked = RGB(240, 240, 240)
End Property

Public Property Get Theme_Error()
    Theme_Error = RGB(255, 205, 210)
End Property


Public Property Get ThemeInactive()
    'a grey that is the same as the hidden rows color
    ThemeInactive = RGB(231, 230, 230)
End Property



Public Property Get ThemeInput()
    'a yellow color used for cell inputs
    ThemeInput = RGB(255, 242, 204)
End Property

