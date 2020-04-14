VERSION 5.00
Begin VB.Form Form_Rand_Case 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "随机生成布局"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5430
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command_Confirm 
      Caption         =   "确定"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   4230
      Width           =   1335
   End
   Begin VB.CommandButton Command_Create 
      Caption         =   "生成"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text_Step 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3080
      Width           =   1335
   End
   Begin VB.TextBox Text_Code 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2560
      Width           =   1335
   End
   Begin VB.Frame Frame 
      Caption         =   "难度"
      Height          =   2295
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Option_Difficulty_5 
         Caption         =   "骨灰"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton Option_Difficulty_4 
         Caption         =   "困难"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton Option_Difficulty_3 
         Caption         =   "中阶"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton Option_Difficulty_2 
         Caption         =   "进阶"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option_Difficulty_1 
         Caption         =   "入门"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form_Rand_Case"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Case_Block
  address As Integer
  style As Integer
End Type
Dim Block(0 To 9) As Case_Block
Dim Rand_Cases(1 To 8000) As String
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Private Sub Form_Load()
  start_x = 150
  start_y = 180
  square_width = 815
  gap = 75
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  Call Get_Rand_Data
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
End Sub
Private Sub Command_Confirm_Click()
  If Option_Difficulty_1.Value = False And Option_Difficulty_2.Value = False And Option_Difficulty_3.Value = False And Option_Difficulty_4.Value = False And Option_Difficulty_5.Value = False Then Exit Sub
  change_case_title = "随机 - "
  If Option_Difficulty_1.Value = True Then change_case_title = change_case_title & "入门"
  If Option_Difficulty_2.Value = True Then change_case_title = change_case_title & "进阶"
  If Option_Difficulty_3.Value = True Then change_case_title = change_case_title & "中阶"
  If Option_Difficulty_4.Value = True Then change_case_title = change_case_title & "困难"
  If Option_Difficulty_5.Value = True Then change_case_title = change_case_title & "骨灰"
  change_case_code = Text_Code
  change_case = True
  Unload Form_Rand_Case
End Sub
Private Sub Command_Create_Click()
  Dim min_step As Integer, max_step As Integer
  Dim index As Long, code As String, step As Integer
  If Option_Difficulty_1.Value = False And Option_Difficulty_2.Value = False And Option_Difficulty_3.Value = False And Option_Difficulty_4.Value = False And Option_Difficulty_5.Value = False Then Exit Sub
  If Option_Difficulty_1.Value = True Then min_step = 0: max_step = 20
  If Option_Difficulty_2.Value = True Then min_step = 21: max_step = 50
  If Option_Difficulty_3.Value = True Then min_step = 51: max_step = 80
  If Option_Difficulty_4.Value = True Then min_step = 81: max_step = 100
  If Option_Difficulty_5.Value = True Then min_step = 101: max_step = 138
  Randomize
retry:
  index = Int(Rnd * 8000) + 1
  code = Left(Rand_Cases(index), 7)
  step = Right(Rand_Cases(index), Len(Rand_Cases(index)) - 9)
  If step < min_step Or step > max_step Then GoTo retry
  Text_Code = code
  Text_Step = step & "步"
  Call Analyse_Code(code)
  Call Output_Graph
End Sub
Private Sub Get_Rand_Data()
  Dim i As Long
  Dim temp As String
  Open "Rand_Cases.txt" For Input As #1
    Do Until EOF(1)
      Line Input #1, temp
      i = i + 1
      Rand_Cases(i) = temp
    Loop
  Close #1
End Sub
Private Sub Output_Graph()
  Dim m, x, y As Integer
  Dim width As Integer, height As Integer
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  For m = 0 To 9
    If Block(m).address <> 25 Then
      x = (Block(m).address Mod 4) * (square_width + gap) + gap + start_x
      y = Int(Block(m).address / 4) * (square_width + gap) + gap + start_y
      If Block(m).style = 0 Or Block(m).style = 1 Then
        width = square_width * 2 + gap
      Else
        width = square_width
      End If
      If Block(m).style = 0 Or Block(m).style = 2 Then
        height = square_width * 2 + gap
      Else
        height = square_width
      End If
      Print_Block x, y, width, height, block_line_width, block_color, block_line_color
    End If
  Next m
End Sub
Private Sub Print_Block(print_start_x, print_start_y, print_width, print_height, print_line_width, print_color, print_line_color)
  If print_width < 0 Or print_height < 0 Then Exit Sub
  FillStyle = 0
  DrawWidth = print_line_width
  FillColor = print_color
  Line (print_start_x, print_start_y)-(print_start_x + print_width, print_start_y + print_height), print_color, B
  Line (print_start_x, print_start_y)-(print_start_x + print_width, print_start_y + print_height), print_line_color, B
End Sub
Private Sub Analyse_Code(code As String)
  On Error Resume Next
  Dim temp(1 To 12) As Integer
  Dim i, addr, style As Integer
  Dim type_1, type_2, type_3 As Integer
  Dim Table(0 To 19) As Integer
  Dim num As Integer, b1 As Integer, b2 As Integer
  Dim dat As String
  For i = 1 To 6
    dat = Mid(code, i + 1, 1)
    If Asc(dat) >= 48 And Asc(dat) <= 57 Then num = Int(dat)
    If Asc(dat) >= 65 And Asc(dat) <= 70 Then num = Asc(dat) - 55
    b1 = num Mod 4
    b2 = (num - b1) / 4 Mod 4
    temp(i * 2 - 1) = b2
    temp(i * 2) = b1
  Next i
  type_1 = 0: type_2 = 0: type_3 = 5
  For i = 0 To 19
    Table(i) = 69
  Next i
  For i = 0 To 9
    Block(i).address = 69
    Block(i).style = 69
  Next i
  dat = Left(code, 1)
  If Asc(dat) >= 48 And Asc(dat) <= 57 Then num = Int(dat)
  If Asc(dat) >= 65 And Asc(dat) <= 70 Then num = Asc(dat) - 55
  Block(0).address = num
  Block(0).style = 0
  If Block(0).address > 14 Then GoTo err
  Table(Block(0).address) = 0
  Table(Block(0).address + 1) = 0
  Table(Block(0).address + 4) = 0
  Table(Block(0).address + 5) = 0
  addr = 0
  For i = 1 To 11
    Do While Table(addr) <> 69
      If addr < 19 Then
        addr = addr + 1
      Else
        Exit Do
      End If
    Loop
    style = temp(i)
    If style = 0 Then
      Table(addr) = 10
    ElseIf style = 1 Then
      If type_2 < 5 Then type_2 = type_2 + 1
      If addr > 18 Then GoTo err
      Block(type_2).style = 1
      Block(type_2).address = addr
      Table(addr) = type_2
      Table(addr + 1) = type_2
    ElseIf style = 2 Then
      If type_2 < 5 Then type_2 = type_2 + 1
      If addr > 15 Then GoTo err
      Block(type_2).style = 2
      Block(type_2).address = addr
      Table(addr) = type_2
      Table(addr + 4) = type_2
    ElseIf style = 3 Then
      If type_3 < 9 Then type_3 = type_3 + 1
      Block(type_3).style = 3
      Block(type_3).address = addr
      Table(addr) = type_3
    End If
  Next i
err:
End Sub
