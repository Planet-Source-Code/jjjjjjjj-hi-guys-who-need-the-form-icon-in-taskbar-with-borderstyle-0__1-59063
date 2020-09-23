Attribute VB_Name = "mdlBorder"
'=============================================================
'            [ Auther : Jim Jose              ]
'            [ Email  : jimjosev33@yahoo.com  ]
'=============================================================
'Hi,
'This code is made for all my friends in PSC. I uploaded this
'code inorder to get useful for anyone. If you found it useful
'please inform me. Your +Ve comments are my motive. Good Luck!
'=============================================================
Option Explicit

'[APIs]
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'[This function can Set your form as BorderStyle=0 ]
'=============================================================
Public Sub SetZeroBorder(Frm As Form)
Dim hRgn As Long
Dim fScaleMode As Long
Dim ScrX As Long, ScrY As Long
Dim fLeft As Long, fTop As Long
Dim fBottom As Long, fRight As Long
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    With Frm
        fScaleMode = .ScaleMode
        .ScaleMode = 1
        fLeft = (.Width - .ScaleWidth) / 2 / ScrX
        fTop = (.Height - .ScaleHeight) / ScrY - fLeft
        fRight = .Width / ScrX - fLeft
        fBottom = .Height / ScrY - fLeft
        hRgn = CreateRectRgn(fLeft, fTop, fRight, fBottom)
        SetWindowRgn .hWnd, hRgn, True
        .ScaleMode = fScaleMode
        DeleteObject hRgn
    End With
End Sub
