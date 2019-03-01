VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Test"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "从内存画到窗体"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "从屏幕画到窗体"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Dim data() As Byte
Dim a As New clsMemDC

Private Sub Command1_Click()
    Me.Cls
    a.CreateMemDC 1920, 1080
    a.BitBltFrom GetDC(0), 0, 0, 0, 0, 1920, 1080
    a.BitBltTo Me.hDC, 0, 0, 0, 0, 1920, 1080
    ReDim data(CLng(1920) * 1080 * 16 / 8)
    a.CopyDataTo data
    a.DeleteMemDC
End Sub

Private Sub Command2_Click()
    Me.Cls
    a.CreateMemDC 1920, 1080
    a.CopyDataFrom data
    a.BitBltTo Me.hDC, 0, 0, 0, 0, 1920, 1080
    a.DeleteMemDC
End Sub
