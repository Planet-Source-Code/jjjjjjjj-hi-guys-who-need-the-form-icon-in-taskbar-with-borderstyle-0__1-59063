VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Sample Form"
   ClientHeight    =   6150
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   10605
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmTest.frx":030A
   ScaleHeight     =   25.625
   ScaleMode       =   4  'Character
   ScaleWidth      =   88.375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00FCE4DA&
      Caption         =   "Set Zero Border"
      Height          =   735
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSet_Click()
    SetZeroBorder Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("( Please 'RATE' this code ).Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( Please give your feedback ) to improve my code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If
    If X = vbOK Then Clipboard.SetText ("Not set")
End Sub
