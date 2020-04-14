VERSION 5.00
Begin VB.Form Form_Favourite 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "我的收藏"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6765
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text_Code 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command_Confirm 
      Caption         =   "确定"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command_Delete 
      Caption         =   "删除"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command_Modify 
      Caption         =   "修改"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command_Add 
      Caption         =   "添加"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.ListBox List_Favourite 
      Height          =   3300
      ItemData        =   "Form_Favourite.frx":0000
      Left            =   3720
      List            =   "Form_Favourite.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form_Favourite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Case_Block
  address As Integer
  style As Integer
End Type
Dim change_mode As Boolean
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
  Call Get_Data
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  If Not List_Favourite.ListCount = 0 Then List_Favourite.ListIndex = 0
End Sub
Private Sub Command_Confirm_Click()
  Dim temp As String
  If List_Favourite.ListCount = 0 Then Exit Sub
  temp = List_Favourite.List(List_Favourite.ListIndex)
  change_case_title = Left(temp, Len(temp) - 9)
  change_case_code = Text_Code
  change_case = True
  Unload Form_Favourite
End Sub
Private Sub Command_Add_Click()
  change_mode = False
  favourite_add_save = False
  favourite_add_init_name = ""
  favourite_add_init_code = ""
  Form_Favourite_Add.Show 1
End Sub
Private Sub Command_Modify_Click()
  Dim temp As String
  If List_Favourite.ListCount = 0 Then Exit Sub
  change_mode = True
  favourite_add_save = False
  temp = List_Favourite.List(List_Favourite.ListIndex)
  favourite_add_init_name = Left(temp, Len(temp) - 9)
  favourite_add_init_code = Left(Right(temp, 8), 7)
  Form_Favourite_Add.Show 1
End Sub
Private Sub Command_Delete_Click()
  Dim temp As Integer
  If List_Favourite.ListCount = 0 Then Exit Sub
  temp = List_Favourite.ListIndex
  List_Favourite.RemoveItem temp
  If List_Favourite.ListCount = temp Then
    List_Favourite.ListIndex = List_Favourite.ListCount - 1
  Else
    List_Favourite.ListIndex = temp
  End If
  If List_Favourite.ListCount = 0 Then
    Text_Code = ""
    Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
  End If
  Call Save_Data
End Sub
Private Sub List_Favourite_Click()
  Dim temp As String
  temp = List_Favourite.List(List_Favourite.ListIndex)
  Text_Code = Mid(temp, Len(temp) - 7, 7)
  Call Analyse_Code(Text_Code)
  Call Output_Graph
End Sub
Private Sub Timer_Timer()
  If favourite_add_confirm = True Then
    If change_mode = True Then Call Command_Delete_Click
    If List_Favourite.ListCount = 0 Then
      List_Favourite.AddItem favourite_add_name & "(" & favourite_add_code & ")"
      List_Favourite.ListIndex = 0
    Else
      List_Favourite.AddItem favourite_add_name & "(" & favourite_add_code & ")", List_Favourite.ListIndex
      List_Favourite.ListIndex = List_Favourite.ListIndex - 1
    End If
    favourite_add_confirm = False
    Call Save_Data
  End If
End Sub
Private Sub Get_Data()
  Dim i As Long
  Call Get_Favourite_Cases
  For i = 1 To UBound(Favourite_Cases_name)
    List_Favourite.AddItem Favourite_Cases_name(i) & "(" & Favourite_Cases_code(i) & ")"
  Next i
End Sub
Private Sub Save_Data()
  Dim i As Integer, temp As String
  ReDim Favourite_Cases_code(List_Favourite.ListCount)
  ReDim Favourite_Cases_name(List_Favourite.ListCount)
  For i = 0 To List_Favourite.ListCount - 1
    temp = List_Favourite.List(i)
    Favourite_Cases_code(i + 1) = Left(Right(temp, 8), 7)
    Favourite_Cases_name(i + 1) = Left(temp, Len(temp) - 9)
  Next i
  Call Save_Favourite_Cases
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
