VERSION 5.00
Begin VB.Form WINnotify 
   Appearance      =   0  'Flat
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
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
      UseMnemonic     =   0   'False
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
    WINnotifyLoaded = True
    SetWindowPos hWnd, -1, 0, 0, 0, 0, 1 Or 2
    Left = Screen.Width - Width - 240
    Top = Screen.Height - Height
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblPostpone(0).BackColor = cButton_Normal
    lblPostpone(0).ForeColor = vbGray
    lblPostpone(1).BackColor = cButton_Normal
    lblPostpone(1).ForeColor = vbGray
    lblPostpone(2).BackColor = cButton_Normal
    lblPostpone(2).ForeColor = vbGray
    lblPostpone(3).BackColor = cButton_Normal
    lblPostpone(3).ForeColor = vbGray
    lblPostpone(4).BackColor = cButton_Normal
    lblPostpone(4).ForeColor = vbGray
    lblPostpone(5).BackColor = cButton_Normal
    lblPostpone(5).ForeColor = vbGray
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WINnotifyLoaded = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WINnotifyLoaded = False
End Sub

Private Sub lblContent_Click()
    On Error Resume Next
    Form_Click
End Sub

Private Sub lblContent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove 0, 0, 0, 0
End Sub

Private Sub lblPostpone_Click(Index As Integer)
    On Error Resume Next
    If Index = 5 Then
        WINremindME.MarkDone 0, Val(Tag)
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
        Dim TempSubject As String
        Dim TempContent As String
        Dim TempTime As String
        Dim TempoTime As String
        Dim TempFreq As String
        Dim TempExpiry As String
        TempSubject = GetSetting("Upcoming", Tag, "Subject", "")
        TempContent = GetSetting("Upcoming", Tag, "Content", "")
        TempoTime = GetSetting("Upcoming", Tag, "oTime", "")
        TempTime = Str(Now + drtn)
        TempFreq = GetSetting("Upcoming", Tag, "Freq", "o")
        TempExpiry = GetSetting("Upcoming", Tag, "Expiry", Str(#12/31/2099#))
        WINremindME.Delete_Work 0, Val(Tag)
        WINremindME.AddTask TempSubject, TempContent, TempTime, TempFreq, TempExpiry, IIf(TempFreq = "o", TempTime, TempoTime)
        If TabSelected = 0 Then
            WINremindME.Clear_All
            WINremindME.Load_All "Upcoming"
        End If
    End If
    Unload Me
End Sub

Private Sub lblPostpone_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblPostpone(Index).BackColor = cButton_Hover
    lblPostpone(Index).ForeColor = vbBlack
End Sub

Private Sub lblSubject_Click()
    On Error Resume Next
    Form_Click
End Sub

Private Sub lblSubject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Form_MouseMove 0, 0, 0, 0
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
