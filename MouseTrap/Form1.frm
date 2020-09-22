VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Release the Mouse Ahhhhhhh!"
      Height          =   690
      Left            =   1350
      TabIndex        =   1
      Top             =   1500
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mouse Trap"
      Height          =   375
      Left            =   1335
      TabIndex        =   0
      Top             =   990
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this code traps your mouse
' it is simple and easy to use

Private Type RECT
    Lft As Long
    Top As Long
    Rght As Long
    Bttom As Long
End Type

Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Public Sub FreeTheMouse(Form)
    Dim erg As Long
    Dim NewRect As RECT


    With NewRect
        .Lft = 0&
        .Top = 0&
        .Rght = Screen.Width / Screen.TwipsPerPixelX
        .Bttom = Screen.Height / Screen.TwipsPerPixelY
    End With
    erg& = ClipCursor(NewRect)
End Sub
Public Sub TrapMouseActivated(Form)
    Dim x As Long, Y As Long, erg As Long
    Dim NewRect As RECT
    x& = Screen.TwipsPerPixelX
    Y& = Screen.TwipsPerPixelY


    With NewRect
        .Lft = Form.Left / x&
        .Top = Form.Top / Y&
        .Rght = .Lft + Form.Width / x&
        .Bttom = .Top + Form.Height / Y&
    End With
    erg& = ClipCursor(NewRect)
End Sub


Private Sub Command1_Click()
Call TrapMouseActivated(Me)
End Sub

Private Sub Command2_Click()
Call FreeTheMouse(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FreeTheMouse(Me)
End Sub
