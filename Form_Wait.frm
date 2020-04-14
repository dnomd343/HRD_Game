VERSION 5.00
Begin VB.Form Form_Wait 
   BorderStyle     =   0  'None
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer_Debug 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text_Debug 
      Appearance      =   0  'Flat
      Height          =   650
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1920
   End
   Begin VB.Timer Timer 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label_Wait 
      AutoSize        =   -1  'True
      Caption         =   " 请稍等哦... "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "Form_Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  If debug_mode = True Then
    Form_Wait.height = 1875
    Text_Debug.Visible = True
  Else
    Form_Wait.height = 465
    Text_Debug.Visible = False
  End If
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  wait_cancel = False
  waiting = True
End Sub
Private Sub Label_Wait_DblClick()
  If MsgBox("真要取消么?", vbYesNo, "> _ <") = vbNo Then Exit Sub
  wait_cancel = True
  waiting = False
  Unload Form_Wait
End Sub
Private Sub Timer_Timer()
  On Error Resume Next
  If Dir(wait_file_name) <> "" Then
    wait_cancel = False
    waiting = False
    Unload Form_Wait
  End If
End Sub
Private Sub Timer_Debug_Timer()
  Dim debug_dat As String
  debug_dat = "wait_cancel=" & wait_cancel & vbCrLf
  debug_dat = debug_dat & "wait_file_name" & vbCrLf & "=" & wait_file_name & vbCrLf
  Text_Debug = debug_dat
End Sub

