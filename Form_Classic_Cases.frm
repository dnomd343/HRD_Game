VERSION 5.00
Begin VB.Form Form_Classic_Cases 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择华容道经典布局"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6990
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command_Confirm 
      Caption         =   "确认"
      Height          =   540
      Left            =   5480
      TabIndex        =   6
      Top             =   4780
      Width           =   1400
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
      Height          =   570
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4780
      Width           =   2175
   End
   Begin VB.CommandButton Command_Search 
      Caption         =   "搜索"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text_Search 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text_Tip 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4780
      Width           =   2895
   End
   Begin VB.ComboBox Combo_Cases 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox List_Cases 
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form_Classic_Cases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Case_Block
  address As Integer
  style As Integer
End Type
Dim tip As String
Dim Block(0 To 9) As Case_Block
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Private Sub Form_Load()
  start_x = 3200
  start_y = 135
  square_width = 815
  gap = 75
  Call Get_Cases_title
  Combo_Cases.ListIndex = 0
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
End Sub
Private Sub Command_Confirm_Click()
  change_case = True
  change_case_title = Left(List_Cases.Text, Len(List_Cases.Text) - 9)
  change_case_code = Text_Code
  Unload Form_Classic_Cases
End Sub
Private Sub List_Cases_Click()
  Dim temp As String
  Text_Tip = "(" & List_Cases.ListIndex + 1 & "/" & List_Cases.ListCount & ")"
  temp = List_Cases.List(List_Cases.ListIndex)
  Text_Code = Mid(temp, Len(temp) - 7, 7)
  Call Analyse_Code(Text_Code)
  Call Output_Graph
End Sub
Private Sub Command_Search_Click()
  Dim i As Integer, j As Integer, last_select As Integer
  Dim temp() As String
  Dim searching As Boolean
  ReDim temp(0)
  If Text_Search = "" Then Exit Sub
  last_select = Combo_Cases.ListIndex
  searching = False
  If Combo_Cases.List(Combo_Cases.ListCount - 1) = "搜索结果" Then
    Combo_Cases.RemoveItem Combo_Cases.ListCount - 1
    searching = True
  End If
  For j = 0 To Combo_Cases.ListCount - 1
    Combo_Cases.ListIndex = j
    If Combo_Cases.Text = "搜索结果" Then Exit For
    For i = 0 To List_Cases.ListCount - 1
      If InStr(List_Cases.List(i), Text_Search) <> 0 Then
        ReDim Preserve temp(UBound(temp) + 1)
        temp(UBound(temp)) = List_Cases.List(i)
      End If
    Next i
  Next j
  List_Cases.Clear
  Combo_Cases.AddItem "搜索结果"
  Combo_Cases.ListIndex = Combo_Cases.ListCount - 1
  Text_Tip = "共找到" & UBound(temp) & "个结果"
  Text_Code = "": Cls
  If UBound(temp) = 0 Then
    If searching = False Then
      Combo_Cases.RemoveItem Combo_Cases.ListCount - 1
      Combo_Cases.ListIndex = last_select
    End If
    MsgBox "No Result!"
    Exit Sub
  End If
  For i = 1 To UBound(temp)
    List_Cases.AddItem temp(i)
  Next i
  List_Cases.ListIndex = 0
  Text_Tip = "共找到" & UBound(temp) & "个结果"
End Sub
Private Sub Combo_Cases_Click()
  If Not Combo_Cases.Text = "搜索结果" Then
    If Combo_Cases.List(Combo_Cases.ListCount - 1) = "搜索结果" Then
      Combo_Cases.RemoveItem Combo_Cases.ListCount - 1
      Text_Search = ""
    End If
    Call Get_Cases(Combo_Cases.ListIndex)
    List_Cases.ListIndex = 0
    Text_Tip = tip
  End If
End Sub
Private Sub Text_Search_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Command_Search_Click
End Sub
Private Sub Get_Cases(index As Integer)
  Dim temp As String, name As String, num As Integer
  num = 0
  List_Cases.Clear
  Open "Classic_Cases.txt" For Input As #1
    Do Until EOF(1)
      Line Input #1, temp
      If temp = "[Cases]" Then
        If num = index Then
        Line Input #1, temp
        Line Input #1, temp
        tip = Right(temp, Len(temp) - 4)
        Text_Tip = tip
reinput:
          If EOF(1) = False Then
            Line Input #1, temp
            If temp <> "[Cases]" Then
              List_Cases.AddItem Right(temp, Len(temp) - 8) & "(" & Left(temp, 7) & ")"
              GoTo reinput
            End If
          End If
        End If
        num = num + 1
      End If
    Loop
  Close #1
End Sub
Private Sub Get_Cases_title()
  Dim temp As String
  Open "Classic_Cases.txt" For Input As #1
    Do Until EOF(1)
      Line Input #1, temp
      If temp = "[Cases]" Then
        Line Input #1, temp
        Combo_Cases.AddItem Right(temp, Len(temp) - 6)
      End If
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

