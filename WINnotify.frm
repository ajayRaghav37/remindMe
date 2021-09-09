VERSION 5.00
Begin VB.Form WINnotify 
   BackColor       =   &H001778A9&
   BorderStyle     =   0  'None
   Caption         =   "remindME Notification"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ShowTimer 
      Interval        =   1
      Left            =   4800
      Top             =   2040
   End
   Begin VB.Label lblPostpone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   5
      Left            =   4605
      TabIndex        =   7
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblPostpone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+1 d"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   4
      Left            =   3705
      TabIndex        =   6
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblPostpone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+12 h"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   3
      Left            =   2805
      TabIndex        =   5
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblPostpone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+6 h"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   2
      Left            =   1905
      TabIndex        =   4
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblPostpone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+1 h"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   1
      Left            =   1005
      TabIndex        =   3
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblPostpone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+15 m"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label lblContent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5010
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4980
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "WINnotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const OneHr As Double = 4.16666666666667E-02
Const FifteenMin As Double = 1.04166666666667E-02
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim pretimer As Double
Dim temp As Integer

Private Sub Form_Click()
    On Error Resume Next
    If Tag <> "" Then
        SaveSetting "Upcoming", Tag, "Clicked", "1"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    SetWindowPos hWnd, -1, 0, 0, 0, 0, 1 Or 2
    Left = Screen.Width - Width - 240
    Top = Screen.Height - Height
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    For temp = 0 To 5
        lblPostpone(temp).BackColor = &H1778A9
        lblPostpone(temp).ForeColor = &H404040
    Next
End Sub

Private Sub lblContent_Click()
    On Error Resume Next
    Call Form_Click
End Sub

Private Sub lblContent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub lblPostpone_Click(Index As Integer)
    On Error Resume Next
    If Index = 5 Then
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Subject", GetSetting("Upcoming", Tag, "Subject", "")
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Content", GetSetting("Upcoming", Tag, "Content", "")
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Time", GetSetting("Upcoming", Tag, "Time", "")
        SaveSetting "Archived", "Total", "Number", Trim(Str(Val(GetSetting("Archived", "Total", "Number", "0")) + 1))
        DeleteSetting "Upcoming", Tag
        Dim WorksLoaded As Integer
        WorksLoaded = Val(GetSetting("Upcoming", "Total", "Number", "0"))
        If WorksLoaded > 1 Then
            For temp = Val(Tag) To WorksLoaded - 2
                SaveSetting "Upcoming", Trim(Str(temp)), "Subject", GetSetting("Upcoming", Trim(Str(temp + 1)), "Subject", "")
                SaveSetting "Upcoming", Trim(Str(temp)), "Time", GetSetting("Upcoming", Trim(Str(temp + 1)), "Time", "")
                SaveSetting "Upcoming", Trim(Str(temp)), "Content", GetSetting("Upcoming", Trim(Str(temp + 1)), "Content", "")
            Next
        End If
        SaveSetting "Upcoming", GetSetting("Upcoming", "Total", "Number", "0"), "Temp", ""
        DeleteSetting "Upcoming", GetSetting("Upcoming", "Total", "Number", "0")
        SaveSetting "Upcoming", "Total", "Number", Trim(Str(Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1))
        If WINremindME.TabSelected = 0 Then
            WINremindME.Clear_All
            WINremindME.Load_All ("Upcoming")
        End If
    Else
        Dim drtn As Double
        Select Case Index
            Case 0
                drtn = FifteenMin
            Case 1
                drtn = OneHr
            Case 2
                drtn = 0.25
            Case 3
                drtn = 0.5
            Case 4
                drtn = 1
        End Select
        SaveSetting "Upcoming", Tag, "Time", Str(Now + drtn)
        If WINremindME.TabSelected = 0 Then
            WINremindME.Load_All ("Upcoming")
        End If
    End If
    Unload Me
End Sub

Private Sub lblPostpone_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblPostpone(Index).BackColor = RGB(225, 165, 38)
    lblPostpone(Index).ForeColor = vbBlack
End Sub

Private Sub lblSubject_Click()
    On Error Resume Next
    Call Form_Click
End Sub

Private Sub lblSubject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub ShowTimer_Timer()
    On Error Resume Next
    ShowTimer.Enabled = False
    If Tag = "" Then
        For temp = 0 To 5
            lblPostpone(temp).Visible = False
        Next
    Else
        For temp = 0 To 5
            lblPostpone(temp).Visible = True
            lblPostpone(temp).ZOrder 0
        Next
    End If
    Do
        Top = Top - 15
        pretimer = Timer
        Do
        Loop Until Timer > pretimer + 0.02
        DoEvents
    Loop Until Top < Screen.Height - Height - 780
End Sub
