VERSION 5.00
Begin VB.Form Form_Creator 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自定义华容道布局"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   5655
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command_Confirm 
      Caption         =   "确定"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command_Clear 
      Caption         =   "清除"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton Command_Mirror 
      Caption         =   "镜像"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text_Code 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton Command_Get_Code 
      Caption         =   "生成编码"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command_Print 
      Caption         =   "解析编码"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Timer Timer_Debug 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text_Debug 
      Height          =   7760
      Left            =   5650
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   180
      Width           =   3495
   End
End
Attribute VB_Name = "Form_Creator"
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
Dim Exist(1 To 4, 1 To 5) As Boolean
Dim Block_index(1 To 4, 1 To 5) As Integer
Dim print_now As Boolean
Dim click_x As Integer, click_y As Integer
Dim click_block_x As Integer, click_block_y As Integer
Dim block_start_x As Integer, block_start_y As Integer, block_width As Integer, block_height As Integer
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Dim x_split(0 To 4) As Integer, y_split(0 To 5) As Integer
Private Sub Form_Load()
  If debug_mode = True Then
    Form_Creator.width = 9400
    Text_Debug.Visible = True
  Else
    Form_Creator.width = 5745
    Text_Debug.Visible = False
  End If
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  Call init
  Call mark
End Sub
Private Sub Form_DblClick()
  Cls
  Call mark
  Call Output_Graph
End Sub
Private Sub init()
  Cls
  start_x = 200
  start_y = 200
  square_width = 1170
  gap = 120
  Call Case_init
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
Private Sub Command_Clear_Click()
  Call Case_init
  Call init
  Cls
  Call mark
End Sub
Private Sub Command_Confirm_Click()
  change_case = True
  change_case_title = "自定义"
  change_case_code = Text_Code
  Unload Form_Creator
End Sub
Private Sub Command_Get_Code_Click()
  If Check_Compete = False Then MsgBox "UnFinished": Exit Sub
  Text_Code = Get_Code
End Sub
Private Sub Command_Mirror_Click()
  Dim i As Integer, addr As Integer
  Dim temp As Integer, temp_b As Boolean
  For i = 0 To 9
    addr = Block(i).address
    If Not addr = 25 Then
      If Block(i).style = 0 Or Block(i).style = 1 Then
        If addr Mod 4 = 0 Then
          addr = addr + 2
        ElseIf addr Mod 4 = 2 Then
          addr = addr - 2
        End If
      End If
      If Block(i).style = 2 Or Block(i).style = 3 Then
        If addr Mod 4 = 0 Then
          addr = addr + 3
        ElseIf addr Mod 4 = 1 Then
          addr = addr + 1
        ElseIf addr Mod 4 = 2 Then
          addr = addr - 1
        ElseIf addr Mod 4 = 3 Then
          addr = addr - 3
        End If
      End If
      Block(i).address = addr
    End If
  Next i
  For i = 1 To 5
    temp_b = Exist(1, i): Exist(1, i) = Exist(4, i): Exist(4, i) = temp_b
    temp_b = Exist(2, i): Exist(2, i) = Exist(3, i): Exist(3, i) = temp_b
    temp = Block_index(1, i): Block_index(1, i) = Block_index(4, i): Block_index(4, i) = temp
    temp = Block_index(2, i): Block_index(2, i) = Block_index(3, i): Block_index(3, i) = temp
  Next i
  If Check_Compete = True Then Text_Code = Get_Code
  Cls
  Call Output_Graph
End Sub
Private Sub Command_Print_Click()
  If Text_Code = "UnFinished" Then
    MsgBox "UnFinished"
  Else
    Text_Code = UCase(Text_Code)
    Analyse (Text_Code)
    If Check = True Then
      Call Output_Graph
    Else
      MsgBox "Error Code!"
      Call Command_Clear_Click
    End If
  End If
End Sub
Private Sub Text_Code_Change()
  If Text_Code = "UnFinished" Then Exit Sub
  If Len(Text_Code) = 7 Then
    Analyse (UCase(Text_Code))
    If Check = True Then
      Call Output_Graph
      Text_Code = UCase(Text_Code)
    Else
      Call Command_Clear_Click
    End If
  Else
    Call Command_Clear_Click
  End If
End Sub
Private Sub Text_Code_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Command_Print_Click
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 And print_now = False Then
    click_x = x
    click_y = y
    print_now = False
    If click_x > start_x And click_x < start_x + square_width * 4 + gap * 5 Then
      If click_y > start_y And click_y < start_y + square_width * 5 + gap * 6 Then
        print_now = True
      End If
    End If
    If print_now = True Then
      click_block_x = Get_block_x(click_x)
      click_block_y = Get_block_y(click_y)
      If Exist(click_block_x, click_block_y) = True Then print_now = False
    End If
    Call Form_MouseMove(Button, Shift, x + 1, y + 1)
  ElseIf Button = 2 Then
    Dim m As Integer
    m = Block_index(Get_block_x(Int(x)), Get_block_y(Int(y)))
    If m <> 10 Then Call Clear_Block(m): Text_Code = "UnFinished"
  End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim m As Integer, addr As Integer
  If Button = 1 And print_now = True Then
    addr = (block_start_y - 1) * 4 + block_start_x - 1
    If block_width = 2 And block_height = 2 Then
      If Block(0).address = 25 Then
        Block(0).address = addr
        Block(0).style = 0
        Exist(block_start_x, block_start_y) = True
        Exist(block_start_x, block_start_y + 1) = True
        Exist(block_start_x + 1, block_start_y) = True
        Exist(block_start_x + 1, block_start_y + 1) = True
        Block_index(block_start_x, block_start_y) = 0
        Block_index(block_start_x, block_start_y + 1) = 0
        Block_index(block_start_x + 1, block_start_y) = 0
        Block_index(block_start_x + 1, block_start_y + 1) = 0
      End If
    End If
    If block_width = 2 And block_height = 1 Then
      For m = 1 To 5
        If Block(m).address = 25 Then
          Block(m).address = addr
          Block(m).style = 1
          Exist(block_start_x, block_start_y) = True
          Exist(block_start_x + 1, block_start_y) = True
          Block_index(block_start_x, block_start_y) = m
          Block_index(block_start_x + 1, block_start_y) = m
          Exit For
        End If
      Next m
    End If
    If block_width = 1 And block_height = 2 Then
      For m = 1 To 5
        If Block(m).address = 25 Then
          Block(m).address = addr
          Block(m).style = 2
          Exist(block_start_x, block_start_y) = True
          Exist(block_start_x, block_start_y + 1) = True
          Block_index(block_start_x, block_start_y) = m
          Block_index(block_start_x, block_start_y + 1) = m
          Exit For
        End If
      Next m
    End If
    If block_width = 1 And block_height = 1 Then
      For m = 6 To 9
        If Block(m).address = 25 Then
          Block(m).address = addr
          Block(m).style = 3
          Exist(block_start_x, block_start_y) = True
          Block_index(block_start_x, block_start_y) = m
          Exit For
        End If
      Next m
    End If
    If Check_Compete = True Then Call Command_Get_Code_Click Else Text_Code = "UnFinished"
  End If
  Call Output_Graph
  print_now = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim output_x As Integer, output_y As Integer, output_width As Integer, output_height As Integer, locked_x As Integer, locked_y As Integer
  If Button = 1 And print_now = True Then
    Call Output_Graph
    If x >= click_x Then
      output_x = start_x + click_block_x * gap + (click_block_x - 1) * square_width
      output_width = x - output_x
      locked_x = x
      If x > start_x + square_width * 4 + gap * 4 Then locked_x = start_x + square_width * 4 + gap * 4: output_width = locked_x - output_x
      If output_width > square_width * 2 + gap Then output_width = square_width * 2 + gap: locked_x = output_x + output_width
      block_start_x = click_block_x
      block_width = Get_block_x(locked_x) - block_start_x + 1
    End If
    If x < click_x Then
      output_x = x
      output_width = (start_x + click_block_x * gap + click_block_x * square_width) - x
      locked_x = x
      If x < start_x + gap Then output_width = (click_block_x - 1) * gap + click_block_x * square_width: locked_x = start_x + gap: output_x = locked_x
      If output_width > square_width * 2 + gap Then locked_x = start_x + (click_block_x - 1) * gap + (click_block_x - 2) * square_width: output_width = square_width * 2 + gap: output_x = locked_x
      block_start_x = Get_block_x(locked_x)
      block_width = Get_block_x(click_x) - block_start_x + 1
    End If
    If y >= click_y Then
      output_y = start_y + click_block_y * gap + (click_block_y - 1) * square_width
      output_height = y - output_y
      locked_y = y
      If y > start_y + square_width * 5 + gap * 5 Then locked_y = start_y + square_width * 5 + gap * 5: output_height = locked_y - output_y
      If output_height > square_width * 2 + gap Then output_height = square_width * 2 + gap: locked_y = output_y + output_height
      block_start_y = click_block_y
      block_height = Get_block_y(locked_y) - block_start_y + 1
    End If
    If y < click_y Then
      output_y = y
      output_height = (start_y + click_block_y * gap + click_block_y * square_width) - y
      locked_y = y
      If y < start_y + gap Then output_height = (click_block_y - 1) * gap + click_block_y * square_width: locked_y = start_y + gap: output_y = locked_y
      If output_height > square_width * 2 + gap Then locked_y = start_y + (click_block_y - 1) * gap + (click_block_y - 2) * square_width: output_height = square_width * 2 + gap: output_y = locked_y
      block_start_y = Get_block_y(locked_y)
      block_height = Get_block_y(click_y) - block_start_y + 1
    End If

    Dim x_limit As Boolean, y_limit As Boolean, xy_limit As Boolean
    If x >= click_x And y >= click_y Then
      x_limit = False: y_limit = False: xy_limit = False
      If block_start_x < 4 Then
        If Exist(block_start_x + 1, block_start_y) = True Then x_limit = True
      Else
        x_limit = True
      End If
      If block_start_y < 5 Then
        If Exist(block_start_x, block_start_y + 1) = True Then y_limit = True
      Else
        y_limit = True
      End If
      If block_start_x < 4 And block_start_y < 5 Then
        If Exist(block_start_x + 1, block_start_y + 1) = True Then xy_limit = True
      End If
      If x_limit = True Then
        If output_width > square_width Then output_width = square_width
        If block_width = 2 Then block_width = 1
      End If
      If y_limit = True Then
        If output_height > square_width Then output_height = square_width
        If block_height = 2 Then block_height = 1
      End If
      If xy_limit = True And x_limit = False And y_limit = False Then
        If output_width < output_height Then
          If output_width > square_width Then output_width = square_width
          If block_width = 2 Then block_width = 1
        Else
          If output_height > square_width Then output_height = square_width
          If block_height = 2 Then block_height = 1
        End If
      End If
    End If
    
    If x >= click_x And y < click_y Then
      x_limit = False: y_limit = False: xy_limit = False
      If block_start_x < 4 Then
        If Exist(block_start_x + 1, block_start_y + block_height - 1) = True Then x_limit = True
      Else
        x_limit = True
      End If
      If block_start_y + block_height - 1 > 1 Then
        If Exist(block_start_x, block_start_y + block_height - 2) = True Then y_limit = True
      Else
        y_limit = True
      End If
      If block_start_x < 4 And block_start_y + block_height - 1 > 1 Then
        If Exist(block_start_x + 1, block_start_y + block_height - 2) = True Then xy_limit = True
      End If
      If x_limit = True Then
        If output_width > square_width Then output_width = square_width
        If block_width = 2 Then block_width = 1
      End If
      If y_limit = True Then
        If output_height > square_width Then
          output_y = output_y + output_height - square_width
          output_height = square_width
        End If
        If block_height = 2 Then block_height = 1: block_start_y = block_start_y + 1
      End If
      If xy_limit = True And x_limit = False And y_limit = False Then
        If output_width < output_height Then
          If output_width > square_width Then output_width = square_width
          If block_width = 2 Then block_width = 1
        Else
          If output_height > square_width Then output_y = output_y + output_height - square_width: output_height = square_width
          If block_height = 2 Then block_height = 1: block_start_y = block_start_y + 1
        End If
      End If
    End If
    
    If x < click_x And y >= click_y Then
      x_limit = False: y_limit = False: xy_limit = False
      If block_start_x + block_width - 1 > 1 Then
        If Exist(block_start_x + block_width - 2, block_start_y) = True Then x_limit = True
      Else
        x_limit = True
      End If
      If block_start_y < 5 Then
        If Exist(block_start_x + block_width - 1, block_start_y + 1) = True Then y_limit = True
      Else
        y_limit = True
      End If
      If block_start_x + block_width - 1 > 1 And block_start_y < 5 Then
        If Exist(block_start_x + block_width - 2, block_start_y + 1) = True Then xy_limit = True
      End If
      If x_limit = True Then
        If output_width > square_width Then
          output_x = output_x + output_width - square_width
          output_width = square_width
        End If
        If block_width = 2 Then block_width = 1: block_start_x = block_start_x + 1
      End If
      If y_limit = True Then
        If output_height > square_width Then output_height = square_width
        If block_height = 2 Then block_height = 1
      End If
      If xy_limit = True And x_limit = False And y_limit = False Then
        If output_width < output_height Then
          If output_width > square_width Then output_x = output_x + output_width - square_width: output_width = square_width
          If block_width = 2 Then block_width = 1: block_start_x = block_start_x + 1
        Else
          If output_height > square_width Then output_height = square_width
          If block_height = 2 Then block_height = 1
        End If
      End If
    End If
    
    If x < click_x And y < click_y Then
      x_limit = False: y_limit = False: xy_limit = False
      If block_start_x + block_width - 1 > 1 Then
        If Exist(block_start_x + block_width - 2, block_start_y + block_height - 1) = True Then x_limit = True
      Else
        x_limit = True
      End If
      If block_start_y + block_height - 1 > 1 Then
        If Exist(block_start_x + block_width - 1, block_start_y + block_height - 2) = True Then y_limit = True
      Else
        y_limit = True
      End If
      If block_start_x + block_width - 1 > 1 And block_start_y < 5 Then
        If Exist(block_start_x + block_width - 2, block_start_y + block_height - 2) = True Then xy_limit = True
      End If
      If x_limit = True Then
        If output_width > square_width Then
          output_x = output_x + output_width - square_width
          output_width = square_width
        End If
        If block_width = 2 Then block_width = 1: block_start_x = block_start_x + 1
      End If
      If y_limit = True Then
        If output_height > square_width Then
          output_y = output_y + output_height - square_width
          output_height = square_width
        End If
        If block_height = 2 Then block_height = 1: block_start_y = block_start_y + 1
      End If
      If xy_limit = True And x_limit = False And y_limit = False Then
        If output_width < output_height Then
          If output_width > square_width Then output_x = output_x + output_width - square_width: output_width = square_width
          If block_width = 2 Then block_width = 1: block_start_x = block_start_x + 1
        Else
          If output_height > square_width Then
            output_y = output_y + output_height - square_width
            output_height = square_width
          End If
          If block_height = 2 Then block_height = 1: block_start_y = block_start_y + 1
        End If
      End If
    End If
    Print_Block output_x, output_y, output_width, output_height, block_line_width, block_color, block_line_color
  End If
End Sub
Private Function Get_block_x(x As Integer) As Integer
  Dim i As Integer
  For i = 1 To 4
    If x > x_split(i - 1) And x < x_split(i) Then
      Get_block_x = i
      Exit For
    End If
  Next i
End Function
Private Function Get_block_y(y As Integer) As Integer
  Dim i As Integer
  For i = 1 To 5
    If y > y_split(i - 1) And y < y_split(i) Then
      Get_block_y = i
      Exit For
    End If
  Next i
End Function
Private Sub mark()
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  If debug_mode = True Then
    Dim i As Integer, j As Integer
    DrawWidth = 1
    For i = 1 To 3
      Line (start_x + gap / 2 + (square_width + gap) * i, start_y)-(start_x + gap / 2 + (square_width + gap) * i, start_y + square_width * 5 + gap * 6)
    Next i
    For i = 1 To 4
      Line (start_x, start_y + gap / 2 + (square_width + gap) * i)-(start_x + square_width * 4 + gap * 5, start_y + gap / 2 + (square_width + gap) * i)
    Next i
    For i = 0 To 3
      For j = 0 To 4
        Line (start_x + square_width * i + gap * (i + 1), start_y + square_width * j + gap * (j + 1))-(start_x + square_width * (i + 1) + gap * (i + 1), start_y + square_width * (j + 1) + gap * (j + 1)), , B
      Next j
    Next i
  End If
End Sub
Private Sub Output_Graph()
  Dim m, x, y As Integer
  Dim width As Integer, height As Integer
  Call mark
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
Private Function Check_Compete()
  Dim m As Integer
  For m = 0 To 9
    If Block(m).style = 4 Then
      Check_Compete = False
      Exit Function
    End If
  Next m
  Check_Compete = True
End Function
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
Private Function Get_Code() As String
  On Error Resume Next
  Dim temp(20) As Boolean
  Dim Table(20) As Integer
  Dim dat(1 To 12) As Integer
  Dim Code As String
  Dim i As Integer, addr As Integer, style As Integer, num As Integer
  For i = 0 To 19
    temp(i) = False
    Table(i) = 10
  Next i
  For i = 0 To 9
    If Block(i).style = 0 Then
      Table(Block(i).address) = i
      Table(Block(i).address + 1) = i
      Table(Block(i).address + 4) = i
      Table(Block(i).address + 5) = i
    ElseIf Block(i).style = 1 Then
      Table(Block(i).address) = i
      Table(Block(i).address + 1) = i
    ElseIf Block(i).style = 2 Then
      Table(Block(i).address) = i
      Table(Block(i).address + 4) = i
    ElseIf Block(i).style = 3 Then
      Table(Block(i).address) = i
    End If
  Next i
  temp(Block(0).address) = True
  temp(Block(0).address + 1) = True
  temp(Block(0).address + 4) = True
  temp(Block(0).address + 5) = True
  If Block(0).address < 10 Then
    Code = Code & Block(0).address
  Else
    Code = Code & Chr(Block(0).address + 55)
  End If
  addr = 0
  num = 1
  For i = 1 To 11
    While (temp(addr) = True)
      If addr < 19 Then
        addr = addr + 1
      Else
        Exit Function
      End If
    Wend
    If Table(addr) = 10 Then
      temp(addr) = True
      dat(num) = 0: num = num + 1
    Else
      style = Block(Table(addr)).style
      If style = 1 Then
        temp(addr) = True
        temp(addr + 1) = True
        dat(num) = 1: num = num + 1
      ElseIf style = 2 Then
        temp(addr) = True
        temp(addr + 4) = True
        dat(num) = 2: num = num + 1
      ElseIf style = 3 Then
        temp(addr) = True
        dat(num) = 3: num = num + 1
      End If
    End If
  Next i
  For i = 1 To 6
    num = dat(i * 2 - 1) * 4 + dat(i * 2)
    If num < 10 Then
      Code = Code & num
    Else
      Code = Code & Chr(num + 55)
    End If
  Next i
  Get_Code = Code
End Function
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
Private Sub Timer_Debug_Timer()
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
  debug_dat = debug_dat & "print_now = " & print_now & vbCrLf & "debug_mode = " & debug_mode
  debug_dat = debug_dat & vbCrLf & vbCrLf
  debug_dat = debug_dat & "click_x = " & click_x & vbCrLf & "click_y = " & click_y & vbCrLf
  debug_dat = debug_dat & "click_block_x = " & click_block_x & vbCrLf & "click_block_y = " & click_block_y & vbCrLf
  debug_dat = debug_dat & "block_start_x = " & block_start_x & vbCrLf & "block_start_y = " & block_start_y & vbCrLf & "block_width = " & block_width & vbCrLf & "block_height = " & block_height
  debug_dat = debug_dat & vbCrLf & vbCrLf
  debug_dat = debug_dat & "start_x = " & start_x & vbCrLf & "start_y = " & start_y & vbCrLf & "gap = " & gap & vbCrLf & "square_width = " & square_width
  debug_dat = debug_dat & vbCrLf & "x_split: "
  For m = 0 To 4
    debug_dat = debug_dat & x_split(m)
    If m <> 4 Then debug_dat = debug_dat & "|"
  Next m
  debug_dat = debug_dat & vbCrLf & "y_split: "
  For m = 0 To 5
    debug_dat = debug_dat & y_split(m)
    If m <> 5 Then debug_dat = debug_dat & "|"
  Next m
  Text_Debug = debug_dat
End Sub

