VERSION 5.00
Begin VB.Form Form_Solution 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "最少步解法"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5295
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer_Play 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command_Last 
      Caption         =   "＞︱"
      Height          =   470
      Left            =   3000
      TabIndex        =   5
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command_Next 
      Caption         =   "＞"
      Height          =   470
      Left            =   2400
      TabIndex        =   4
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command_Pause 
      Caption         =   "播放"
      Height          =   470
      Left            =   1320
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command_Previous 
      Caption         =   "＜"
      Height          =   470
      Left            =   720
      TabIndex        =   2
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command_First 
      Caption         =   "︱＜"
      Height          =   470
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox List_Solution 
      Height          =   4740
      ItemData        =   "Form_Solution.frx":0000
      Left            =   3720
      List            =   "Form_Solution.frx":0002
      TabIndex        =   0
      Top             =   290
      Width           =   1455
   End
   Begin VB.Timer Timer_Get_Data 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label_Index 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   0
      TabIndex        =   6
      Top             =   80
      Width           =   90
   End
End
Attribute VB_Name = "Form_Solution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Case_Block
  address As Integer
  style As Integer
End Type
Dim wait_data As Boolean
Dim Block(0 To 9) As Case_Block
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Private Sub Form_Load()
  start_x = 135
  start_y = 135
  square_width = 770
  gap = 75
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  wait_file_name = start_code & ".txt"
  If Dir(start_code & ".txt") <> "" Then Kill start_code & ".txt"
  Shell "Engine.exe -q " & start_code
  wait_cancel = False
  waiting = True
  wait_data = True
  Form_Wait.Show 1
End Sub
Private Sub Command_First_Click()
  List_Solution.ListIndex = 0
End Sub
Private Sub Command_Last_Click()
  List_Solution.ListIndex = List_Solution.ListCount - 1
End Sub
Private Sub Command_Previous_Click()
  If List_Solution.ListIndex > 0 Then
    List_Solution.ListIndex = List_Solution.ListIndex - 1
  End If
End Sub
Private Sub Command_Next_Click()
  If List_Solution.ListIndex < List_Solution.ListCount - 1 Then
    List_Solution.ListIndex = List_Solution.ListIndex + 1
  End If
End Sub
Private Sub Command_Pause_Click()
  If Timer_Play.Enabled = False Then
    Command_Pause.Caption = "暂停"
    Timer_Play.Enabled = True
  Else
    Command_Pause.Caption = "播放"
    Timer_Play.Enabled = False
  End If
End Sub
Private Sub List_Solution_Click()
  If Not Label_Index = "无解" Then Label_Index = "(" & List_Solution.ListIndex & "/" & List_Solution.ListCount - 1 & ")"
  Label_Index.Left = List_Solution.Left + (List_Solution.width - Label_Index.width) / 2
  Call Analyse_Code(List_Solution.List(List_Solution.ListIndex))
  Call Output_Graph
End Sub
Private Sub Timer_Play_Timer()
  If List_Solution.ListIndex = List_Solution.ListCount - 1 Then
    Command_Pause.Caption = "播放"
    Timer_Play.Enabled = False
  End If
  Call Command_Next_Click
End Sub
Private Sub Timer_Get_Data_Timer()
  If wait_data = True And waiting = False Then
    wait_data = False
    Call Get_Data(start_code & ".txt")
  End If
End Sub
Private Sub Get_Data(file_name As String)
  Dim temp As String, i As Integer, num As Integer
  Open file_name For Input As #1
    Line Input #1, temp
    If temp = "No Solution" Then
      MsgBox "无解啊啊啊", , "> _ <"
      Label_Index.Caption = "无解"
      List_Solution.AddItem start_code
    Else
      num = Int(temp)
      MsgBox "只要" & num & "步就行啦", , "> _ <"
      For i = 0 To num
        Line Input #1, temp
        List_Solution.AddItem temp
      Next i
    End If
  Close #1
  List_Solution.ListIndex = 0
End Sub
Private Sub Output_Graph()
  Dim m, X, Y As Integer
  Dim width As Integer, height As Integer
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  For m = 0 To 9
    If Block(m).address <> 25 Then
      X = (Block(m).address Mod 4) * (square_width + gap) + gap + start_x
      Y = Int(Block(m).address / 4) * (square_width + gap) + gap + start_y
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
      Print_Block X, Y, width, height, block_line_width, block_color, block_line_color
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
