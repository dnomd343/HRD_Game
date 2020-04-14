VERSION 5.00
Begin VB.Form Form_Detail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "详细信息"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7965
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command_Analyse 
      Caption         =   "全局溯源分析"
      Height          =   300
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer_Debug 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text_Debug 
      Height          =   4380
      Left            =   7960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Timer Timer_Get_Data 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox List_Data 
      Height          =   4020
      ItemData        =   "Form_Detail.frx":0000
      Left            =   2520
      List            =   "Form_Detail.frx":0002
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox Combo_Detail 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox List_Detail 
      Height          =   4020
      ItemData        =   "Form_Detail.frx":0004
      Left            =   120
      List            =   "Form_Detail.frx":0006
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Case_Block
  address As Integer
  style As Integer
End Type
Private Type Layer_struct
  size As Integer
  layer_dat() As String
End Type
Dim wait_data As Boolean, loading As Boolean
Dim Block(0 To 9) As Case_Block
Dim start_x As Integer, start_y As Integer, square_width As Integer, gap As Integer
Dim group_size As Long, min_steps As Integer, farthest_steps As Integer
Dim min_solutions() As String, farthest_cases() As String, solutions() As String, layers() As String, layer() As Layer_struct
Private Sub Form_Load()
  start_x = 4350
  start_y = 135
  square_width = 777
  gap = 75
  loading = False
  If debug_mode = True Then
    Form_Detail.width = 10575
    Text_Debug.Visible = True
  Else
    Form_Detail.width = 8055
    Text_Debug.Visible = False
  End If
  If on_top = True Then
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 1 Or 2
  Else
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 1 Or 2
  End If
  ReDim min_solutions(0)
  ReDim farthest_cases(0)
  ReDim solutions(0)
  ReDim layers(0)
  ReDim layer(0 To 0)
  Combo_Detail.AddItem "最少步解"
  Combo_Detail.AddItem "所有的解"
  Combo_Detail.AddItem "最远的布局"
  Combo_Detail.AddItem "各步数的布局"
  wait_file_name = start_code & ".txt"
  If Dir(start_code & ".txt") <> "" Then Kill start_code & ".txt"
  Shell "Engine.exe -a " & start_code
  wait_cancel = False
  waiting = True
  wait_data = True
  Form_Wait.Show 1
  Print_Block start_x, start_y, square_width * 4 + gap * 5, square_width * 5 + gap * 6, case_line_width, case_color, case_line_color
End Sub
Private Sub Combo_Detail_Click()
  Dim i As Long
  List_Detail.Clear
  If Combo_Detail.ListIndex = 0 Then
    If min_steps = -1 Then
      List_Detail.AddItem "无解"
    Else
      List_Detail.AddItem Combo_Detail.Text & "(" & min_steps & "步,共" & UBound(min_solutions) & "个)"
    End If
  ElseIf Combo_Detail.ListIndex = 1 Then
    List_Detail.AddItem Combo_Detail.Text & "(共" & UBound(solutions) & "个)"
  ElseIf Combo_Detail.ListIndex = 2 Then
    List_Detail.AddItem Combo_Detail.Text & "(" & farthest_steps & "步,共" & UBound(farthest_cases) & "个)"
  ElseIf Combo_Detail.ListIndex = 3 Then
    For i = 0 To UBound(layer)
      List_Detail.AddItem "第" & i & "步(共" & layer(i).size & "个)"
    Next i
  End If
  List_Detail.ListIndex = 0
End Sub
Private Sub List_Detail_Click()
  Dim i As Long, n As Integer
  loading = True
  List_Data.Clear
  If Combo_Detail.ListIndex = 0 Then
    For i = 1 To UBound(min_solutions)
      If Not min_steps = -1 Then List_Data.AddItem min_solutions(i) & "(" & min_steps & "步)"
    Next i
  ElseIf Combo_Detail.ListIndex = 1 Then
    For i = 1 To UBound(solutions)
      n = n + 1
      If n = 200 Then n = 0: DoEvents
      List_Data.AddItem Left(solutions(i), 7) & Mid(solutions(i), 9, Len(solutions(i)) - 9) & "步)"
    Next i
  ElseIf Combo_Detail.ListIndex = 2 Then
    For i = 1 To UBound(farthest_cases)
      List_Data.AddItem farthest_cases(i) & "(" & farthest_steps & "步)"
    Next i
  ElseIf Combo_Detail.ListIndex = 3 Then
    For i = 0 To UBound(layer(List_Detail.ListIndex).layer_dat)
      List_Data.AddItem layer(List_Detail.ListIndex).layer_dat(i) & "(" & List_Detail.ListIndex & "步)"
    Next i
  End If
  If Not min_steps = -1 Then
    List_Data.ListIndex = 0
  Else
    If Combo_Detail.ListIndex = 2 Or Combo_Detail.ListIndex = 3 Then List_Data.ListIndex = 0
  End If
  loading = False
End Sub
Private Sub List_Data_Click()
  Call Analyse_Code(Left(List_Data.List(List_Data.ListIndex), 7))
  Call Output_Graph
End Sub
Private Sub Timer_Get_Data_Timer()
  Dim dat As String
  Combo_Detail.Enabled = Not loading
  If wait_data = True And waiting = False Then
    wait_data = False
    If wait_cancel = True Then
      Unload Form_Detail
    Else
      MsgBox Form_Game.Label_Title, , "> _ <"
      Call Get_Data(start_code & ".txt")
      dat = "共衍生出" & group_size & "种布局" & vbCrLf & "最远为" & farthest_steps & "步" & vbCrLf
      If min_steps = -1 Then dat = dat & "无解" Else dat = dat & "最少需要" & min_steps & "步"
      MsgBox dat, , "> _ <"
      Combo_Detail.ListIndex = 0
    End If
  End If
End Sub
Private Sub Command_Analyse_Click()
  MsgBox "还没做好呢QAQ", , "> _ <"
End Sub
Private Sub Get_Data(file_name As String)
  Dim temp As String
  ReDim min_solutions(0)
  ReDim farthest_cases(0)
  ReDim solutions(0)
  ReDim layers(0)
  Open file_name For Input As #1
    Line Input #1, temp: Line Input #1, temp
    group_size = temp
    Line Input #1, temp: Line Input #1, temp
    min_steps = temp
    Line Input #1, temp: Line Input #1, temp
    farthest_steps = temp
    Line Input #1, temp: Line Input #1, temp
    While (temp <> "[Farthest_cases]")
      ReDim Preserve min_solutions(UBound(min_solutions) + 1)
      min_solutions(UBound(min_solutions)) = temp
      Line Input #1, temp
    Wend
    Line Input #1, temp
    While (temp <> "[Solutions]")
      ReDim Preserve farthest_cases(UBound(farthest_cases) + 1)
      farthest_cases(UBound(farthest_cases)) = temp
      Line Input #1, temp
    Wend
    Line Input #1, temp
    While (temp <> "[List]")
      ReDim Preserve solutions(UBound(solutions) + 1)
      solutions(UBound(solutions)) = temp
      Line Input #1, temp
    Wend
    Line Input #1, temp
    While (temp <> "[Layer]")
      ReDim Preserve layers(UBound(layers) + 1)
      layers(UBound(layers)) = temp
      Line Input #1, temp
    Wend
  Close #1
  Call split_layer
End Sub
Private Sub split_layer()
  Dim i As Long, code As String, num As Integer, index As Integer
  For i = 1 To UBound(layers)
    code = Mid(layers(i), InStr(1, layers(i), ">") + 2, 7)
    num = Mid(layers(i), InStr(1, layers(i), "(") + 1, InStr(1, layers(i), ",") - InStr(1, layers(i), "(") - 1)
    index = Mid(layers(i), InStr(1, layers(i), ",") + 1, Len(layers(i)) - InStr(1, layers(i), ",") - 1)
    ReDim Preserve layer(0 To num)
    ReDim Preserve layer(num).layer_dat(0 To index)
    layer(num).layer_dat(index) = code
    layer(num).size = index + 1
  Next i
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
Private Sub Timer_Debug_Timer()
  Dim debug_dat As String
  debug_dat = debug_dat & "group_size=" & group_size & vbCrLf
  debug_dat = debug_dat & "min_steps=" & min_steps & vbCrLf
  debug_dat = debug_dat & "farthest_steps=" & farthest_steps & vbCrLf
  debug_dat = debug_dat & vbCrLf
  debug_dat = debug_dat & "min_solutions->" & UBound(min_solutions) & vbCrLf
  debug_dat = debug_dat & "farthest_cases->" & UBound(farthest_cases) & vbCrLf
  debug_dat = debug_dat & "solutions->" & UBound(solutions) & vbCrLf
  debug_dat = debug_dat & "layers->" & UBound(layers) & vbCrLf
  debug_dat = debug_dat & "layer->" & UBound(layer) & vbCrLf
  Text_Debug = debug_dat
End Sub
