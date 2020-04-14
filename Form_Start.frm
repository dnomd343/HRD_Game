VERSION 5.00
Begin VB.Form Form_Start 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择初始布局"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3585
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2300
      Width           =   180
   End
   Begin VB.CommandButton Command_Favourite 
      Caption         =   "收藏的布局"
      Height          =   1095
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command_Rand_Case 
      Caption         =   "随机生成布局"
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command_Select_Case 
      Caption         =   "选择经典布局"
      Height          =   1095
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command_Create_Case 
      Caption         =   "自定义布局"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form_Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
End Sub
Private Sub Command_Create_Case_Click()
  Form_Creator.Show 1
End Sub
Private Sub Command_Select_Case_Click()
  Form_Classic_Cases.Show 1
End Sub
Private Sub Command_Rand_Case_Click()
  Form_Rand_Case.Show 1
End Sub
Private Sub Command_Favourite_Click()
  favourite_add_confirm = False
  Form_Favourite.Show 1
End Sub
Private Sub Timer_Timer()
  If change_case = True Then
    Form_Game.Show
    Unload Form_Start
  End If
End Sub
