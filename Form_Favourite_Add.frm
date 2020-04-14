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
   Begin VB.CommandButton Command_Confirm 
      Caption         =   "确认"
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   4800
      Width           =   1120
   End
   Begin VB.TextBox Text_Code 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   600
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text_Name 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label_Code 
      AutoSize        =   -1  'True
      Caption         =   "编码:"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   5200
      Width           =   450
   End
   Begin VB.Label Label_Name 
      AutoSize        =   -1  'True
      Caption         =   "名称:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   4840
      Width           =   450
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
Dim Rand_Cases(1 To 8000) As String
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Private Sub Form_Load()
  start_x = 120
  start_y = 120
  square_width = 815
  gap = 75
  favourite_add_confirm = False
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  Text_Name = favourite_add_init_name
  Text_Code = favourite_add_init_code
  Call Text_Code_Change
End Sub
Private Sub Command_Confirm_Click()
  If Text_Name = "" Then MsgBox "你还没有填名称喔", , "(⊙-⊙)": Exit Sub
  Call Analyse_Code(UCase(Text_Code))
  If Check = False Then MsgBox "编码出错啦", , "(⊙-⊙)": Exit Sub
  favourite_add_confirm = True
  favourite_add_name = Text_Name
  favourite_add_code = Text_Code
  Unload Form_Favourite_Add
End Sub
Private Sub Text_Code_Change()
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  If Len(Text_Code) = 7 Then
    Call Analyse_Code(UCase(Text_Code))
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
