VERSION 5.00
Begin VB.Form Form_Favourite_Add 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "我的收藏"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   3870
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Label_Code 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Text            =   "编码:"
      Top             =   5200
      Width           =   495
   End
   Begin VB.TextBox Label_Name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Text            =   "名称:"
      Top             =   4840
      Width           =   495
   End
   Begin VB.Timer Timer_Debug 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text_Debug 
      Height          =   5320
      Left            =   3880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   100
      Width           =   3375
   End
   Begin VB.CommandButton Command_Confirm 
      Caption         =   "确认"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   4800
      Width           =   1120
   End
   Begin VB.TextBox Text_Code 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   600
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text_Name 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   600
      TabIndex        =   0
      Top             =   4800
      Width           =   1935
   End
End
Attribute VB_Name = "Form_Favourite_Add"
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
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Dim x_split(0 To 4) As Integer, y_split(0 To 5) As Integer
Dim click_mouse_x As Integer, click_mouse_y As Integer
Dim click_x As Integer, click_y As Integer, mouse_button As Integer, print_now As Boolean
Dim delta_x As Integer, delta_y As Integer, locked_x As Integer, locked_y As Integer
Dim limit(-1 To 1, -1 To 1) As Boolean

Private Sub Form_DblClick()
  If mouse_button = 2 Then
    Call Case_init
    Call Output_Graph
    Text_Code = ""
  End If
End Sub

Private Sub Form_Load()
  start_x = 120
  start_y = 120
  square_width = 815
  gap = 75
  print_now = False
  favourite_add_confirm = False
  If debug_mode = True Then
    Form_Favourite_Add.width = 7425
    Text_Debug.Visible = True
  Else
    Form_Favourite_Add.width = 3960
    Text_Debug.Visible = False
  End If
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  Call init
  Call Case_init
  Text_Name = favourite_add_init_name
  Text_Code = favourite_add_init_code
  Call Text_Code_Change
End Sub
Private Sub Command_Confirm_Click()
  If Text_Name = "" Then MsgBox "你还没有填名称喔", , "(⊙-⊙)": Exit Sub
  Call Analyse(UCase(Text_Code))
  If Check = False Then MsgBox "编码出错啦", , "(⊙-⊙)": Exit Sub
  favourite_add_confirm = True
  favourite_add_name = Text_Name
  favourite_add_code = Text_Code
  If favourite_add_save = True Then
    favourite_add_save = False
    Call Get_Favourite_Cases
    ReDim Preserve Favourite_Cases_code(UBound(Favourite_Cases_code) + 1)
    ReDim Preserve Favourite_Cases_name(UBound(Favourite_Cases_name) + 1)
    Favourite_Cases_code(UBound(Favourite_Cases_code)) = favourite_add_code
    Favourite_Cases_name(UBound(Favourite_Cases_name)) = favourite_add_name
    Call Save_Favourite_Cases
  End If
  Unload Form_Favourite_Add
End Sub


Private Sub Form_Unload(Cancel As Integer)
  favourite_add_save = False
End Sub

Private Sub Label_Name_Click()
  Text_Name.SetFocus
End Sub
Private Sub Label_Code_Click()
  Text_Code.SetFocus
End Sub
Private Sub Text_Code_Change()
  If print_now = True Then Exit Sub
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  If Len(Text_Code) = 7 Then
    Call Analyse(UCase(Text_Code))
    If Check = True Then
      Text_Code = UCase(Text_Code)
      Call Output_Graph
    End If
  End If
End Sub
Private Sub Text_Code_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Command_Confirm_Click
End Sub
Private Sub Text_Name_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Text_Code.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  click_mouse_x = X
  click_mouse_y = Y
  click_x = Get_block_x(Int(X))
  click_y = Get_block_y(Int(Y))
  mouse_button = Button
  If click_x = 0 Or click_x = 5 Then Exit Sub
  If click_y = 0 Or click_y = 6 Then Exit Sub
  If Exist(click_x, click_y) = True Then
    Call Clear_Block(Block_index(click_x, click_y))
    Text_Code = ""
    Call Output_Graph
    Exit Sub
  End If
  If Not Button = 1 Then Exit Sub
  Call check_limit(click_x, click_y)
  print_now = True
  Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim print_x As Integer, print_y As Integer, print_width As Integer, print_height As Integer
  If print_now = True Then
    delta_x = Get_Signed(Get_block_x(Int(X)) - click_x)
    delta_y = Get_Signed(Get_block_y(Int(Y)) - click_y)
    If delta_x = 0 And delta_y = 0 Then
      locked_x = click_x
      locked_y = click_y
    ElseIf Abs(delta_x) = 1 And Abs(delta_y) = 1 Then
      locked_x = click_x + delta_x
      locked_y = click_y + delta_y
      If limit(delta_x, delta_y) = True And limit(delta_x, 0) = False And limit(0, delta_y) = False Then
        If Abs(click_mouse_x - X) < Abs(click_mouse_y - Y) Then locked_x = click_x Else locked_y = click_y
      End If
      If limit(delta_x, 0) = True Then locked_x = click_x
      If limit(0, delta_y) = True Then locked_y = click_y
    ElseIf Abs(delta_x) = 1 And Abs(delta_y) = 0 Then
      locked_y = click_y
      If limit(delta_x, delta_y) = True Then locked_x = click_x Else locked_x = click_x + delta_x
    ElseIf Abs(delta_x) = 0 And Abs(delta_y) = 1 Then
      locked_x = click_x
      If limit(delta_x, delta_y) = True Then locked_y = click_y Else locked_y = click_y + delta_y
    End If
    print_x = Get_Min(click_x, locked_x) * (square_width + gap) - square_width + start_x
    print_y = Get_Min(click_y, locked_y) * (square_width + gap) - square_width + start_y
    If locked_x = click_x Then print_width = square_width Else print_width = square_width * 2 + gap
    If locked_y = click_y Then print_height = square_width Else print_height = square_width * 2 + gap
    Call Output_Graph
    Print_Block print_x, print_y, print_width, print_height, block_line_width, block_color, block_line_color
  End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim block_start_x As Integer, block_start_y As Integer, block_width As Integer, block_height As Integer
  Dim addr As Integer, m As Integer
  If print_now = True Then
    block_start_x = Get_Min(click_x, locked_x)
    block_start_y = Get_Min(click_y, locked_y)
    block_width = Abs(click_x - locked_x) + 1
    block_height = Abs(click_y - locked_y) + 1
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
    Text_Code = ""
    Call Output_Graph
    If Check_Compete = True Then Text_Code = Get_Code
    print_now = False
  End If
End Sub
Private Sub check_limit(X As Integer, Y As Integer)
  Dim i As Integer, j As Integer
  For i = -1 To 1
    For j = -1 To 1
      limit(i, j) = False
    Next j
  Next i
  If X = 1 Then
    limit(-1, -1) = True: limit(-1, 0) = True: limit(-1, 1) = True
  Else
    If Exist(X - 1, Y) = True Then limit(-1, -1) = True: limit(-1, 0) = True: limit(-1, 1) = True
    If Not Y = 1 Then
      If Exist(X - 1, Y - 1) = True Then limit(-1, -1) = True
    End If
    If Not Y = 5 Then
      If Exist(X - 1, Y + 1) = True Then limit(-1, 1) = True
    End If
  End If
  If X = 4 Then
    limit(1, -1) = True: limit(1, 0) = True: limit(1, 1) = True
  Else
    If Exist(X + 1, Y) = True Then limit(1, -1) = True: limit(1, 0) = True: limit(1, 1) = True
    If Not Y = 1 Then
      If Exist(X + 1, Y - 1) = True Then limit(1, -1) = True
    End If
    If Not Y = 5 Then
      If Exist(X + 1, Y + 1) = True Then limit(1, 1) = True
    End If
  End If
  If Y = 1 Then
    limit(-1, -1) = True: limit(0, -1) = True: limit(1, -1) = True
  Else
    If Exist(X, Y - 1) = True Then limit(-1, -1) = True: limit(0, -1) = True: limit(1, -1) = True
    If Not X = 1 Then
      If Exist(X - 1, Y - 1) = True Then limit(-1, -1) = True
    End If
    If Not X = 4 Then
      If Exist(X + 1, Y - 1) = True Then limit(1, -1) = True
    End If
  End If
  If Y = 5 Then
    limit(-1, 1) = True: limit(0, 1) = True: limit(1, 1) = True
  Else
    If Exist(X, Y + 1) = True Then limit(-1, 1) = True: limit(0, 1) = True: limit(1, 1) = True
    If Not X = 1 Then
      If Exist(X - 1, Y + 1) = True Then limit(-1, 1) = True
    End If
    If Not X = 4 Then
      If Exist(X + 1, Y + 1) = True Then limit(1, 1) = True
    End If
  End If
  If Not Block(0).address = 25 Then limit(-1, -1) = True: limit(-1, 1) = True: limit(1, -1) = True: limit(1, 1) = True
End Sub
Private Function Get_Min(num_1 As Integer, num_2 As Integer) As Integer
  If num_1 < num_2 Then Get_Min = num_1 Else Get_Min = num_2
End Function
Private Function Get_Signed(num As Integer) As Integer
  If num > 0 Then Get_Signed = 1
  If num = 0 Then Get_Signed = 0
  If num < 0 Then Get_Signed = -1
End Function
Private Function Get_block_x(X As Integer) As Integer
  Dim i As Integer
  For i = 1 To 4
    If X >= x_split(i - 1) And X <= x_split(i) Then
      Get_block_x = i
      Exit For
    End If
  Next i
  If X < x_split(0) Then Get_block_x = 0
  If X > x_split(4) Then Get_block_x = 5
End Function
Private Function Get_block_y(Y As Integer) As Integer
  Dim i As Integer
  For i = 1 To 5
    If Y >= y_split(i - 1) And Y <= y_split(i) Then
      Get_block_y = i
      Exit For
    End If
  Next i
  If Y < y_split(0) Then Get_block_y = 0
  If Y > y_split(5) Then Get_block_y = 6
End Function
Private Sub init()
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
  For i = -1 To 1
    For j = -1 To 1
      limit(i, j) = False
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
  Dim X As Integer, Y As Integer, addr As Integer, style As Integer
  addr = Block(m).address
  style = Block(m).style
  Block(m).address = 25
  Block(m).style = 4
  Y = Int(addr / 4) + 1
  X = addr - (Y - 1) * 4 + 1
  If style = 0 Then
    Exist(X, Y) = False
    Exist(X, Y + 1) = False
    Exist(X + 1, Y) = False
    Exist(X + 1, Y + 1) = False
    Block_index(X, Y) = 10
    Block_index(X, Y + 1) = 10
    Block_index(X + 1, Y) = 10
    Block_index(X + 1, Y + 1) = 10
  End If
  If style = 1 Then
    Exist(X, Y) = False
    Exist(X + 1, Y) = False
    Block_index(X, Y) = 10
    Block_index(X + 1, Y) = 10
  End If
  If style = 2 Then
    Exist(X, Y) = False
    Exist(X, Y + 1) = False
    Block_index(X, Y) = 10
    Block_index(X, Y + 1) = 10
  End If
  If style = 3 Then
    Exist(X, Y) = False
    Block_index(X, Y) = 10
  End If
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
Private Function Get_Code() As String
  On Error Resume Next
  Dim temp(20) As Boolean
  Dim Table(20) As Integer
  Dim dat(1 To 12) As Integer
  Dim code As String
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
    code = code & Block(0).address
  Else
    code = code & Chr(Block(0).address + 55)
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
      code = code & num
    Else
      code = code & Chr(num + 55)
    End If
  Next i
  Get_Code = code
End Function
Private Sub Analyse(code As String)
  Dim m As Integer, addr As Integer, X As Integer, Y As Integer
  Call Analyse_Code(code)
  For X = 1 To 4
    For Y = 1 To 5
      Block_index(X, Y) = 10
      Exist(X, Y) = False
    Next Y
  Next X
  For m = 0 To 9
    addr = Block(m).address
    Y = Int(addr / 4) + 1
    X = addr - (Y - 1) * 4 + 1
    If Block(m).style = 0 Then
      Block_index(X, Y) = 0
      Block_index(X, Y + 1) = 0
      Block_index(X + 1, Y) = 0
      Block_index(X + 1, Y + 1) = 0
    End If
    If Block(m).style = 1 Then
      Block_index(X, Y) = m
      Block_index(X + 1, Y) = m
    End If
    If Block(m).style = 2 Then
      Block_index(X, Y) = m
      Block_index(X, Y + 1) = m
    End If
    If Block(m).style = 3 Then
      Block_index(X, Y) = m
    End If
  Next m
  For X = 1 To 4
    For Y = 1 To 5
      If Block_index(X, Y) <> 10 Then Exist(X, Y) = True
    Next Y
  Next X
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

Private Sub Timer_Debug_Timer()
  Dim debug_dat As String
  Dim i As Integer, j As Integer, m As Integer
  For m = 0 To 9
    debug_dat = debug_dat & "Block[" & m & "] -> address = " & Block(m).address & "  style = " & Block(m).style
    If m <> 9 Then debug_dat = debug_dat & vbCrLf
  Next m
  debug_dat = debug_dat & vbCrLf & vbCrLf
  debug_dat = debug_dat & "   exist     block_index   limit" & vbCrLf
  For j = 1 To 5
    For i = 1 To 4
      If Exist(i, j) Then
        debug_dat = debug_dat & "$$ "
      Else
        debug_dat = debug_dat & "[] "
      End If
    Next i
    debug_dat = debug_dat & "   "
    For i = 1 To 4
      If Block_index(i, j) = 10 Then
        debug_dat = debug_dat & "A "
      Else
        debug_dat = debug_dat & Trim(Block_index(i, j)) & " "
      End If
    Next i
    debug_dat = debug_dat & "   "
    If j <= 3 Then
      For i = -1 To 1
        If limit(i, j - 2) = True Then
          debug_dat = debug_dat & "$$ "
        Else
          debug_dat = debug_dat & "[] "
        End If
      Next i
    End If
    debug_dat = debug_dat & vbCrLf & vbCrLf
  Next j
  debug_dat = debug_dat & "click_mouse_x=" & click_mouse_x & vbCrLf & "click_mouse_y=" & click_mouse_y & vbCrLf
  debug_dat = debug_dat & "click_x=" & click_x & " " & "click_y=" & click_y & vbCrLf
  debug_dat = debug_dat & "delta_x=" & delta_x & " " & "delta_y=" & delta_y & vbCrLf
  debug_dat = debug_dat & "locked_x=" & locked_x & " " & "locked_y=" & locked_y & vbCrLf
  debug_dat = debug_dat & "print_now=" & print_now & vbCrLf
  debug_dat = debug_dat & "favourite_add_save=" & favourite_add_save & vbCrLf
  Text_Debug = debug_dat
End Sub
