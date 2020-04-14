Attribute VB_Name = "Module"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public debug_mode As Boolean, playing As Boolean, solve_compete As Boolean
Public block_line_width As Integer, case_line_width As Integer
Public block_color, block_line_color, case_color, case_line_color
Public change_case As Boolean, change_case_title As String, change_case_code As String



