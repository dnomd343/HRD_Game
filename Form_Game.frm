VERSION 5.00
Begin VB.Form Form_Game 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HRD Game v0.0 by Dnomd343"
   ClientHeight    =   6990
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   9660
   Icon            =   "Form_Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   9840
      Top             =   240
   End
   Begin VB.TextBox Text_Debug 
      Height          =   6735
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu test 
      Caption         =   "test"
   End
End
Attribute VB_Name = "Form_Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Case_Block
  address As Integer
  style As Integer
End Type
Private Type Block_Address
  x As Integer
  y As Integer
End Type
Dim Block(0 To 9) As Case_Block
Dim Exist(1 To 4, 1 To 5) As Boolean
Dim Block_index(1 To 4, 1 To 5) As Integer
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Dim block_line_width As Integer, case_line_width As Integer
Dim block_color, block_line_color, case_color, case_line_color
Dim x_split(0 To 4) As Integer, y_split(0 To 5) As Integer
Dim dir_x1 As Integer, dir_y1 As Integer, dir_x2 As Integer, dir_y2 As Integer
Dim block_addr(0 To 2) As Block_Address, move_max_step As Integer
Dim mouse_x As Long, mouse_y As Long
Dim last_move As Integer, move_times As Integer
Dim total_steps As Long
Dim debug_mode As Boolean
Private Sub Form_Click()
  Dim m As Integer
  If mouse_x < start_x Or mouse_x > start_x + square_width * 4 + gap * 5 Then Exit Sub
  If mouse_y < start_y Or mouse_y > start_y + square_width * 5 + gap * 6 Then Exit Sub
  m = Block_index(Get_block_x(mouse_x), Get_block_y(mouse_y))
  If m = 10 Then Exit Sub
  If m = last_move Then
    If move_max_step = 1 Then
      If dir_x2 = 0 And dir_y2 = 0 Then
        If move_times Mod 2 = 1 Then
          Call Move_Block(m, -dir_x1, -dir_y1)
        Else
          Call Move_Block(m, dir_x1, dir_y1)
        End If
      Else
        If move_times Mod 4 = 0 Then
          Call Move_Block(m, dir_x1, dir_y1)
        ElseIf move_times Mod 4 = 1 Then
          Call Move_Block(m, block_addr(0).x - block_addr(1).x, block_addr(0).y - block_addr(1).y)
        ElseIf move_times Mod 4 = 2 Then
          Call Move_Block(m, block_addr(2).x - block_addr(0).x, block_addr(2).y - block_addr(0).y)
        Else
          Call Move_Block(m, block_addr(0).x - block_addr(2).x, block_addr(0).y - block_addr(2).y)
        End If
      End If
    ElseIf move_max_step = 2 Then
      If move_times Mod 4 = 0 Then
        Call Move_Block(m, dir_x1, dir_y1)
      ElseIf move_times Mod 4 = 1 Then
        Call Move_Block(m, block_addr(2).x - block_addr(1).x, block_addr(2).y - block_addr(1).y)
      ElseIf move_times Mod 4 = 2 Then
        Call Move_Block(m, block_addr(1).x - block_addr(2).x, block_addr(1).y - block_addr(2).y)
      Else
        Call Move_Block(m, block_addr(0).x - block_addr(1).x, block_addr(0).y - block_addr(1).y)
      End If
    End If
    move_times = move_times + 1
  Else
    Call Check_Move(m)
    move_times = 1
    last_move = m
    If move_max_step = 0 Then Exit Sub
    total_steps = total_steps + 1
    Call Move_Block(m, dir_x1, dir_y1)
  End If
  Call Output_Graph
End Sub

Private Sub Form_DblClick()
  Call Form_Click
End Sub

Private Sub Form_Load()
  Me.Icon = LoadPicture("")
  Call init
  debug_mode = True
  Call Analyse("1A9BF0C")
  Call Output_Graph
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  mouse_x = x
  mouse_y = y
End Sub
Private Sub Move_Block(m As Integer, dir_x As Integer, dir_y As Integer)
  Dim addr As Integer, style As Integer, x As Integer, y As Integer
  addr = Block(m).address
  style = Block(m).style
  y = Int(addr / 4) + 1
  x = addr - (y - 1) * 4 + 1
  x = x + dir_x
  y = y + dir_y
  addr = (y - 1) * 4 + x - 1
  Call Clear_Block(m)
  Block(m).address = addr
  Block(m).style = style
  If Block(m).style = 0 Then
    Block_index(x, y) = m
    Block_index(x, y + 1) = m
    Block_index(x + 1, y) = m
    Block_index(x + 1, y + 1) = m
  End If
  If Block(m).style = 1 Then
    Block_index(x, y) = m
    Block_index(x + 1, y) = m
  End If
  If Block(m).style = 2 Then
    Block_index(x, y) = m
    Block_index(x, y + 1) = m
  End If
  If Block(m).style = 3 Then
    Block_index(x, y) = m
  End If
  For x = 1 To 4
    For y = 1 To 5
      If Block_index(x, y) <> 10 Then Exist(x, y) = True
    Next y
  Next x
End Sub
Private Sub Check_Move(m As Integer)
  Dim addr As Integer, x As Integer, y As Integer
  Dim move_once As Boolean
  move_once = False
  dir_x1 = 0: dir_x2 = 0: dir_y1 = 0: dir_y2 = 0
  move_max_step = 0
  addr = Block(m).address
  y = Int(addr / 4) + 1
  x = addr - (y - 1) * 4 + 1
  block_addr(0).x = x: block_addr(0).y = y
  block_addr(1).x = x: block_addr(1).y = y
  block_addr(2).x = x: block_addr(2).y = y
  If Block(m).style = 0 Then
    If y > 1 Then
      If Exist(x, y - 1) = False And Exist(x + 1, y - 1) = False Then move_max_step = 1: dir_y1 = -1
    End If
    If y < 4 Then
      If Exist(x, y + 2) = False And Exist(x + 1, y + 2) = False Then move_max_step = 1: dir_y1 = 1
    End If
    If x > 1 Then
      If Exist(x - 1, y) = False And Exist(x - 1, y + 1) = False Then move_max_step = 1: dir_x1 = -1
    End If
    If x < 3 Then
      If Exist(x + 2, y) = False And Exist(x + 2, y + 1) = False Then move_max_step = 1: dir_x1 = 1
    End If
  ElseIf Block(m).style = 1 Then
    If y > 1 Then
      If Exist(x, y - 1) = False And Exist(x + 1, y - 1) = False Then move_max_step = 1: dir_y1 = -1
    End If
    If y < 5 Then
      If Exist(x, y + 1) = False And Exist(x + 1, y + 1) = False Then move_max_step = 1: dir_y1 = 1
    End If
    If x > 1 Then
      If Exist(x - 1, y) = False Then
        move_max_step = 1
        If move_once = False Then dir_x1 = -1 Else dir_x2 = -1
        move_once = True
        If x > 2 Then
          If Exist(x - 2, y) = False Then move_max_step = 2: dir_x2 = -2
        End If
      End If
    End If
    If x < 3 Then
      If Exist(x + 2, y) = False Then
        move_max_step = 1
        If move_once = False Then dir_x1 = 1 Else dir_x2 = 1
        move_once = True
        If x < 2 Then
          If Exist(x + 3, y) = False Then move_max_step = 2: dir_x2 = 2
        End If
      End If
    End If
  ElseIf Block(m).style = 2 Then
    If y > 1 Then
      If Exist(x, y - 1) = False Then
        move_max_step = 1
        If move_once = False Then dir_y1 = -1 Else dir_y2 = -1
        move_once = True
        If y > 2 Then
          If Exist(x, y - 2) = False Then move_max_step = 2: dir_y2 = -2
        End If
      End If
    End If
    If y < 4 Then
      If Exist(x, y + 2) = False Then
        move_max_step = 1
        If move_once = False Then dir_y1 = 1 Else dir_y2 = 1
        move_once = True
        If y < 3 Then
          If Exist(x, y + 3) = False Then move_max_step = 2: dir_y2 = 2
        End If
      End If
    End If
    If x > 1 Then
      If Exist(x - 1, y) = False And Exist(x - 1, y + 1) = False Then move_max_step = 1: dir_x1 = -1
    End If
    If x < 4 Then
      If Exist(x + 1, y) = False And Exist(x + 1, y + 1) = False Then move_max_step = 1: dir_x1 = 1
    End If
  ElseIf Block(m).style = 3 Then
    If y > 1 Then
      If Exist(x, y - 1) = False Then
        move_max_step = 1
        If move_once = False Then dir_y1 = -1 Else dir_y2 = -1
        move_once = True
        If y > 2 Then
          If Exist(x, y - 2) = False Then move_max_step = 2: dir_y2 = -2
        End If
        If x > 1 Then
          If Exist(x - 1, y - 1) = False Then move_max_step = 2: dir_x2 = -1: dir_y2 = -1
        End If
        If x < 4 Then
          If Exist(x + 1, y - 1) = False Then move_max_step = 2: dir_x2 = 1: dir_y2 = -1
        End If
      End If
    End If
    If y < 5 Then
      If Exist(x, y + 1) = False Then
        move_max_step = 1
        If move_once = False Then dir_y1 = 1 Else dir_y2 = 1
        move_once = True
        If y < 4 Then
          If Exist(x, y + 2) = False Then move_max_step = 2: dir_y2 = 2
        End If
        If x > 1 Then
          If Exist(x - 1, y + 1) = False Then move_max_step = 2: dir_x2 = -1: dir_y2 = 1
        End If
        If x < 4 Then
          If Exist(x + 1, y + 1) = False Then move_max_step = 2: dir_x2 = 1: dir_y2 = 1
        End If
      End If
    End If
    If x > 1 Then
      If Exist(x - 1, y) = False Then
        move_max_step = 1
        If move_once = False Then dir_x1 = -1 Else dir_x2 = -1
        move_once = True
        If x > 2 Then
          If Exist(x - 2, y) = False Then move_max_step = 2: dir_x2 = -2
        End If
        If y > 1 Then
          If Exist(x - 1, y - 1) = False Then move_max_step = 2: dir_x2 = -1: dir_y2 = -1
        End If
        If y < 5 Then
          If Exist(x - 1, y + 1) = False Then move_max_step = 2: dir_x2 = -1: dir_y2 = 1
        End If
      End If
    End If
    If x < 4 Then
      If Exist(x + 1, y) = False Then
        move_max_step = 1
        If move_once = False Then dir_x1 = 1 Else dir_x2 = 1
        move_once = True
        If x < 3 Then
          If Exist(x + 2, y) = False Then move_max_step = 2: dir_x2 = 2
        End If
        If y > 1 Then
          If Exist(x + 1, y - 1) = False Then move_max_step = 2: dir_x2 = 1: dir_y2 = -1
        End If
        If y < 5 Then
          If Exist(x + 1, y + 1) = False Then move_max_step = 2: dir_x2 = 1: dir_y2 = 1
        End If
      End If
    End If
  End If
  block_addr(1).x = block_addr(0).x + dir_x1
  block_addr(1).y = block_addr(0).y + dir_y1
  block_addr(2).x = block_addr(0).x + dir_x2
  block_addr(2).y = block_addr(0).y + dir_y2
End Sub
Private Function Get_block_x(x As Long) As Integer
  Dim i As Integer
  For i = 1 To 4
    If x > x_split(i - 1) And x < x_split(i) Then
      Get_block_x = i
      Exit For
    End If
  Next i
End Function
Private Function Get_block_y(y As Long) As Integer
  Dim i As Integer
  For i = 1 To 5
    If y > y_split(i - 1) And y < y_split(i) Then
      Get_block_y = i
      Exit For
    End If
  Next i
End Function
Private Sub init()
  last_move = 10
  move_times = 0
  total_steps = 0
  start_x = 90
  start_y = 180
  gap = 105
  square_width = 1200
  block_line_width = 1
  case_line_width = 2
  block_line_color = RGB(0, 0, 0)
  case_line_color = RGB(0, 0, 0)
  block_color = RGB(250, 250, 250)
  case_color = RGB(256, 256, 256)
  x_split(0) = start_x
  x_split(1) = start_x + gap / 2 + square_width + gap
  x_split(2) = start_x + gap / 2 + (square_width + gap) * 2
  x_split(3) = start_x + gap / 2 + (square_width + gap) * 3
  x_split(4) = start_x + gap + (square_width + gap) * 4
  y_split(0) = start_y
  y_split(1) = start_y + gap / 2 + square_width + gap
  y_split(2) = start_y + gap / 2 + (square_width + gap) * 2
  y_split(3) = start_y + gap / 2 + (square_width + gap) * 3
  y_split(4) = start_y + gap / 2 + (square_width + gap) * 4
  y_split(5) = start_y + gap + (square_width + gap) * 5
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
Private Sub Case_init()
  Dim i As Integer, j As Integer
  For i = 0 To 9
    Block(i).address = 25
    Block(i).style = 4
  Next i
  For i = 1 To 4
    For j = 1 To 5
      Exist(i, j) = False
      Block_index(i, j) = 10
    Next j
  Next i
End Sub
Private Sub Clear_Block(m As Integer)
  Dim x As Integer, y As Integer, addr As Integer
  addr = Block(m).address
  y = Int(addr / 4) + 1
  x = addr - (y - 1) * 4 + 1
  If Block(m).style = 0 Then
    Exist(x, y) = False
    Exist(x, y + 1) = False
    Exist(x + 1, y) = False
    Exist(x + 1, y + 1) = False
    Block_index(x, y) = 10
    Block_index(x, y + 1) = 10
    Block_index(x + 1, y) = 10
    Block_index(x + 1, y + 1) = 10
  End If
  If Block(m).style = 1 Then
    Exist(x, y) = False
    Exist(x + 1, y) = False
    Block_index(x, y) = 10
    Block_index(x + 1, y) = 10
  End If
  If Block(m).style = 2 Then
    Exist(x, y) = False
    Exist(x, y + 1) = False
    Block_index(x, y) = 10
    Block_index(x, y + 1) = 10
  End If
  If Block(m).style = 3 Then
    Exist(x, y) = False
    Block_index(x, y) = 10
  End If
  Block(m).address = 25
  Block(m).style = 4
End Sub
Private Sub Analyse(Code As String)
  Dim m As Integer, addr As Integer, x As Integer, y As Integer
  Call Analyse_Code(Code)
  For x = 1 To 4
    For y = 1 To 5
      Block_index(x, y) = 10
      Exist(x, y) = False
    Next y
  Next x
  For m = 0 To 9
    addr = Block(m).address
    y = Int(addr / 4) + 1
    x = addr - (y - 1) * 4 + 1
    If Block(m).style = 0 Then
      Block_index(x, y) = 0
      Block_index(x, y + 1) = 0
      Block_index(x + 1, y) = 0
      Block_index(x + 1, y + 1) = 0
    End If
    If Block(m).style = 1 Then
      Block_index(x, y) = m
      Block_index(x + 1, y) = m
    End If
    If Block(m).style = 2 Then
      Block_index(x, y) = m
      Block_index(x, y + 1) = m
    End If
    If Block(m).style = 3 Then
      Block_index(x, y) = m
    End If
  Next m
  For x = 1 To 4
    For y = 1 To 5
      If Block_index(x, y) <> 10 Then Exist(x, y) = True
    Next y
  Next x
End Sub
Private Function Check() As Boolean
  Dim temp(0 To 19) As Boolean
  Dim addr As Integer, i As Integer, j As Integer
  For i = 0 To 19
    temp(i) = False
  Next i
  Check = True
  If Block(0).style <> 0 Or Block(0).address > 20 Or Block(0).address < 0 Then
    Check = False
  Else
    addr = Block(0).address
    If addr > 14 Or (addr Mod 4 = 3) Then Check = False
    temp(addr) = True
    temp(addr + 1) = True
    temp(addr + 4) = True
    temp(addr + 5) = True
  End If
  For i = 1 To 5
    If Block(i).address > 20 Or Block(i).address < 0 Then
      Check = False
    ElseIf Block(i).style <> 1 And Block(i).style <> 2 Then
      Check = False
    Else
      addr = Block(i).address
      If Block(i).style = 1 Then
        If addr > 18 Or (addr Mod 4 = 3) Then Check = False
        If temp(addr) = True Or temp(addr + 1) = True Then Check = False
        temp(addr) = True
        temp(addr + 1) = True
      End If
      If Block(i).style = 2 Then
        If addr > 15 Then Check = False
        If temp(addr) = True Or temp(addr + 4) = True Then Check = False
        temp(addr) = True
        temp(addr + 4) = True
      End If
    End If
  Next i
  For i = 6 To 9
    If Block(i).style <> 3 Or Block(i).address > 20 Or Block(i).address < 0 Then
      Check = False
    Else
      addr = Block(i).address
      If addr > 19 Then Check = False
      If temp(addr) = True Then Check = False
      temp(addr) = True
    End If
  Next i
  j = 0
  For i = 0 To 19
    If temp(i) = False Then j = j + 1
  Next i
  If j <> 2 Then Check = False
End Function
Private Sub Analyse_Code(Code As String)
  On Error Resume Next
  Dim temp(1 To 12) As Integer
  Dim i, addr, style As Integer
  Dim type_1, type_2, type_3 As Integer
  Dim Table(0 To 19) As Integer
  Dim num As Integer, b1 As Integer, b2 As Integer
  Dim dat As String
  For i = 1 To 6
    dat = Mid(Code, i + 1, 1)
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
  dat = Left(Code, 1)
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
Private Sub Timer1_Timer()
  Dim i As Integer, j As Integer, m As Integer, debug_dat As String
  For m = 0 To 9
    debug_dat = debug_dat & "Block[" & m & "] -> address = " & Block(m).address & "  style = " & Block(m).style
    If m <> 9 Then debug_dat = debug_dat & vbCrLf
  Next m
  debug_dat = debug_dat & vbCrLf & vbCrLf
  debug_dat = debug_dat & "   exist          block_index" & vbCrLf
  For j = 1 To 5
    For i = 1 To 4
      If Exist(i, j) Then
        debug_dat = debug_dat & "$$ "
      Else
        debug_dat = debug_dat & "[] "
      End If
    Next i
    debug_dat = debug_dat & "        "
    For i = 1 To 4
      If Block_index(i, j) = 10 Then
        debug_dat = debug_dat & "A "
      Else
        debug_dat = debug_dat & Trim(Block_index(i, j)) & " "
      End If
    Next i
    debug_dat = debug_dat & vbCrLf & vbCrLf
  Next j
  debug_dat = debug_dat & "dir_x1=" & dir_x1 & " dir_y1=" & dir_y1 & vbCrLf
  debug_dat = debug_dat & "dir_x2=" & dir_x2 & " dir_y2=" & dir_y2 & vbCrLf
  debug_dat = debug_dat & "block_addr(0)=(" & block_addr(0).x & "," & block_addr(0).y & ")" & vbCrLf
  debug_dat = debug_dat & "block_addr(1)=(" & block_addr(1).x & "," & block_addr(1).y & ")" & vbCrLf
  debug_dat = debug_dat & "block_addr(2)=(" & block_addr(2).x & "," & block_addr(2).y & ")" & vbCrLf
  debug_dat = debug_dat & "move_max_step=" & move_max_step & vbCrLf
  debug_dat = debug_dat & "last_move=" & last_move & vbCrLf
  debug_dat = debug_dat & "move_times=" & move_times & vbCrLf
  debug_dat = debug_dat & "total_steps=" & total_steps & vbCrLf
  Text_Debug = debug_dat
End Sub
