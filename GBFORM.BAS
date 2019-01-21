Attribute VB_Name = "GBFORM"

Option Explicit

Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Global Const GWW_HINSTANCE = -6
Global Const GWW_HWNDPARENT = -8
Global Const GWW_ID = -12

Global Const MODELESS = 0
Global Const MODAL = 1

Global Const User = 0        ' 0 - User
Global Const TWIPS = 1       ' 1 - Twip
Global Const POINTS = 2      ' 2 - Point
Global Const PIXELS = 3      ' 3 - Pixel
Global Const CHARACTERS = 4  ' 4 - Character
Global Const INCHES = 5      ' 5 - Inch
Global Const MILLIMETERS = 6 ' 6 - Millimeter
Global Const CENTIMETERS = 7 ' 7 - Centimeter
Sub gbasForm_Centre(frmForm As Form)
    frmForm.Top = (Screen.Height - frmForm.Height) / 2
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
End Sub
Function gbasForm_GetID(frmForm As Form) As Integer
    gbasForm_GetID = GetWindowWord(frmForm.hwnd, GWW_ID)
End Function

Function gbasForm_GetInstance(frmForm As Form) As Integer
    gbasForm_GetInstance = GetWindowWord(frmForm.hwnd, GWW_HINSTANCE)
End Function

Function gbasForm_GetParent(frmForm As Form) As Integer
    gbasForm_GetParent = GetWindowWord(frmForm.hwnd, GWW_HWNDPARENT)
End Function
Function gbasForm_StayOnTop(frmForm As Form, blnOnTop As Integer) As Integer
    Dim intWindow     As Integer
    Dim intPosition   As Integer
    Dim intFlags      As Integer
    Dim intStatus     As Integer
    Dim blnStatus     As Integer
    intWindow = frmForm.hwnd
    If blnOnTop Then
        intPosition = HWND_TOPMOST
    Else
        intPosition = HWND_NOTOPMOST
    End If
    intFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    intStatus = SetWindowPos(intWindow, intPosition, 0, 0, 0, 0, intFlags)
    If intStatus = 0 Then
        blnStatus = False
    Else
        blnStatus = True
    End If
    gbasForm_StayOnTop = blnStatus
End Function

