Attribute VB_Name = "Initializer"
Public Const cScroll_Normal = &H1D719A
Public Const cScroll_Hover = &H3185AE
Public Const cScroll_Pressed = &H41E2B
Public Const cButton_Normal = &H1778A9
Public Const cButton_Hover = &H26A5E1
Public Const cTab_Normal = &H95075
Public Const cTab_Selected = &H125D82
Public Const vbDarkRed = &H96
Public Const vbDarkGray = &H101010
Public Const vbGray = &H404040
Public Const OneHr As Double = 4.16666666666667E-02
Public Const FifteenMin As Double = 1.04166666666667E-02
Public Const OneMin As Double = 6.94444444444444E-04
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public TabSelected As Byte
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public WINnotifyLoaded As Boolean
