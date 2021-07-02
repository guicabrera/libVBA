Attribute VB_Name = "ajustaScreenResize"
Option Explicit

Declare PtrSafe Function FindWindowA& Lib "User32" (ByVal lpClassName$, ByVal lpWindowName$)
Declare PtrSafe Function GetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&)
Declare PtrSafe Function SetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)


' Déclaration des constantes
Public Const GWL_STYLE As Long = -16
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_FULLSIZING = &H70000
Public Lg As Single
Public Ht As Single
Public Fini As Boolean
Dim meValue As Object

'sub para ser chamada no método de initialize do userform
Sub iniVars(ByVal meObj As Object)

    Ht = meObj.Height
    Lg = meObj.Width
    InitMaxMin (meObj.Caption)
    Application.WindowState = xlMaximized
'    meObj.Height = Application.Height
'    meObj.Width = Application.Width
    meObj.Left = Application.Left
    meObj.Top = Application.Top

End Sub

'sub para ser chamada no método de resize do userform
Sub ajusteItensUF(ByVal meObj As Object)
Dim RtL As Single, RtH As Single
        If meObj.Width < 300 Or meObj.Height < 200 Or Fini Then Exit Sub
        RtL = meObj.Width / Lg
        RtH = meObj.Height / Ht
        meObj.Zoom = IIf(RtL < RtH, RtL, RtH) * 100
End Sub

'sub que adiciona a opção de maximizar e minimizar

Sub InitMaxMin(mCaption As String, Optional Max As Boolean = True, Optional Min As Boolean = True _
        , Optional Sizing As Boolean = True)
Dim hWnd As Long
    hWnd = FindWindowA(vbNullString, mCaption)
    If Max Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
    If Min Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MINIMIZEBOX
    If Sizing Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_FULLSIZING
End Sub
