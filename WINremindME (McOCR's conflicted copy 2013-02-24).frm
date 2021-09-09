VERSION 5.00
Begin VB.Form WINremindME 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "remindME"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   Icon            =   "WINremindME.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "WINremindME.frx":89EA
   ScaleHeight     =   6540
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Notify 
      Interval        =   1000
      Left            =   6840
      Top             =   5520
   End
   Begin VB.Frame UpcomingFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00125D82&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   8550
      Begin VB.TextBox txtContent 
         Appearance      =   0  'Flat
         BackColor       =   &H0026A5E1&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.Label lblExpand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00151515&
         Height          =   360
         Left            =   7260
         TabIndex        =   15
         ToolTipText     =   "Expand"
         Top             =   240
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00151515&
         Height          =   360
         Left            =   7590
         TabIndex        =   14
         ToolTipText     =   "Done"
         Top             =   240
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblDelete 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00151515&
         Height          =   270
         Left            =   7920
         TabIndex        =   13
         ToolTipText     =   "Delete"
         Top             =   300
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblSubject 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   300
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblWork 
         BackColor       =   &H001778A9&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.Label lblScroll 
         BackColor       =   &H001D719A&
         Height          =   615
         Left            =   8280
         TabIndex        =   2
         Top             =   120
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Label lblHourly 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4485
      TabIndex        =   21
      ToolTipText     =   $"WINremindME.frx":D8E4C
      Top             =   6180
      Width           =   2550
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Hourly Reminders: Disabled"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   4695
      TabIndex        =   20
      Top             =   6270
      Width           =   2130
   End
   Begin VB.Line LnJunk 
      Index           =   3
      X1              =   4485
      X2              =   4485
      Y1              =   6180
      Y2              =   6555
   End
   Begin VB.Label lblFeedback 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7965
      TabIndex        =   19
      Top             =   6180
      Width           =   1200
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7050
      TabIndex        =   18
      Top             =   6180
      Width           =   900
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A&bout"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   7050
      TabIndex        =   17
      Top             =   6270
      Width           =   900
   End
   Begin VB.Line LnJunk 
      Index           =   2
      X1              =   7035
      X2              =   7035
      Y1              =   6180
      Y2              =   6555
   End
   Begin VB.Line LnJunk 
      Index           =   1
      X1              =   9165
      X2              =   9165
      Y1              =   6180
      Y2              =   6555
   End
   Begin VB.Line LnJunk 
      Index           =   0
      X1              =   7950
      X2              =   7950
      Y1              =   6180
      Y2              =   6555
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Feedback"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   7965
      TabIndex        =   16
      Top             =   6270
      Width           =   1200
   End
   Begin VB.Label BtnClose 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   9330
      TabIndex        =   12
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   240
   End
   Begin VB.Label lblArchived 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Archived"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblPending 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblUpcoming 
      Alignment       =   2  'Center
      BackColor       =   &H00125D82&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2655
      TabIndex        =   8
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   4455
   End
   Begin VB.Shape sTab 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00125D82&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   600
      Top             =   1440
      Width           =   8565
   End
   Begin VB.Shape sTab 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00095075&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   7110
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Shape sTab 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00095075&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   600
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Shape sTab 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00125D82&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   600
      Top             =   4800
      Width           =   8565
   End
   Begin VB.Label cmdAddTask 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Add Task (A)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7440
      TabIndex        =   7
      Top             =   5520
      Width           =   1725
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   655
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   6
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   9735
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Upcoming"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2655
      TabIndex        =   9
      Top             =   1050
      UseMnemonic     =   0   'False
      Width           =   4455
   End
   Begin VB.Shape sTab 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00125D82&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   2655
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "WINremindME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_DLGFRAME = &H400000
Dim Xinit As Long
Dim Yinit As Long
Dim Mousedwn As Boolean
Dim WorksLoaded As Integer
Dim TempNum As Integer
Dim TempNum2 As Integer
Dim TempNum3 As Integer
Dim WorkSelectedTop As Integer
Dim ContentNo As Integer
Dim CalledExt As Boolean
Dim WorkSelected As Integer
Dim JustHide As Boolean
Dim nEvents As Integer
Dim nEvents2 As Integer
Dim HeaderLen As Integer

Private Sub BtnClose_Click()
    On Error Resume Next
    WindowState = 1
End Sub

Private Sub cmdAddTask_Click()
    On Error Resume Next
    cmdAddTask.BackColor = cButton_Normal
    Load WINtask
    WINtask.Left = Left + (Width - WINtask.Width) / 2
    WINtask.Top = Top + (Height - WINtask.Height) / 2
    WINtask.Show vbModal
    AddTask WINtask.txtSubject.Text, WINtask.txtContent.Text, WINtask.txtDeadline.Text, WINtask.cmdOK.Tag, WINtask.txtExpiry.Text
    Unload WINtask
End Sub

Private Sub cmdAddTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdAddTask.BackColor = cButton_Hover
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If WINabout.Visible Then
        Unload WINabout
        JustHide = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim TempNum As Integer
    Select Case KeyCode
        Case vbKeyTab
            Select Case TabSelected
                Case 0
                    If Shift = 0 Then
                        lblPending_Click
                    Else
                        lblArchived_Click
                    End If
                Case 1
                    If Shift = 0 Then
                        lblUpcoming_Click
                    Else
                        lblPending_Click
                    End If
                Case 2
                    If Shift = 0 Then
                        lblArchived_Click
                    Else
                        lblUpcoming_Click
                    End If
            End Select
        Case vbKeyReturn
            If WorksLoaded > 0 And TabSelected <> 1 Then
                Edit_Work
            End If
        Case vbKeyLeft
            If txtContent.Visible Then
                Expand_Content WorkSelected
            End If
        Case vbKeyRight
            If Not txtContent.Visible And WorksLoaded > 0 Then
                Expand_Content WorkSelected
            End If
        Case vbKeyUp
            If WorkSelected < WorksLoaded - 1 And Not txtContent.Visible Then
                lblWork_Click WorkSelected + 1
            End If
        Case vbKeyDown
            If WorkSelected > 0 And Not txtContent.Visible Then
                lblWork_Click WorkSelected - 1
            End If
        Case vbKeyHome
            If Not txtContent.Visible Then
                lblWork_Click WorksLoaded - 1
            End If
        Case vbKeyEnd
            If Not txtContent.Visible Then
                lblWork_Click 0
            End If
        Case vbKeyPageUp
            If Not txtContent.Visible Then
                TempNum = IIf(WorkSelected + 5 < WorksLoaded, WorkSelected + 5, WorksLoaded - 1)
                lblWork_Click TempNum
            End If
        Case vbKeyPageDown
            If Not txtContent.Visible Then
                TempNum = IIf(WorkSelected - 5 >= 0, WorkSelected - 5, 0)
                lblWork_Click TempNum
            End If
        Case vbKeyDelete
            Delete_Work TabSelected, WorkSelected
    End Select
    If WorksLoaded > 0 Then
        lblDelete.Top = lblWork(WorkSelected).Top + 180
        lblDone.Top = lblWork(WorkSelected).Top + 150
        lblExpand.Top = lblWork(WorkSelected).Top + 150
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyA
            Select Case Shift
                Case 0
                    cmdAddTask_Click
                Case vbAltMask
                    cmdAddTask_Click
                Case Is > vbCtrlMask
                    If txtContent.Visible Then
                        txtContent.SelStart = 0
                        txtContent.SelLength = Len(txtContent.Text)
                    End If
                Case vbCtrlMask
                    If txtContent.Visible Then
                        txtContent.SelStart = HeaderLen
                        txtContent.SelLength = Len(txtContent.Text)
                    End If
            End Select
        Case vbKeyB
            If Shift = vbAltMask Then
                lblAbout_Click
            End If
        Case vbKeyC
            If Shift = vbCtrlMask Then
                If txtContent.Visible Then
                    If txtContent.SelLength > 0 Then
                        Clipboard.Clear
                        Clipboard.SetText txtContent.SelText
                    End If
                End If
            End If
        Case vbKeyF
            If Shift = vbAltMask Then
                lblFeedback_Click
            End If
        Case vbKeyH
            If Shift = vbAltMask Then
                lblHourly_Click
            End If
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then
        End
    End If
    Dim ShortcutSystem As Object
    Dim CurrentShortcut As Object
    Set ShortcutSystem = CreateObject("WScript.Shell")
    Set CurrentShortcut = ShortcutSystem.createshortcut(ShortcutSystem.SpecialFolders("Startup") + "\remindME.lnk")
    CurrentShortcut.targetpath = App.Path + "\" + App.EXEName + ".exe"
    CurrentShortcut.workingdirectory = App.Path
    CurrentShortcut.save
    SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) + WS_DLGFRAME
    Show
    Height = 6555
    Width = 9750
    WorkSelected = -1
    Load_All "Upcoming"
    If GetSetting("remindME", "Notify", "Latest Check", "") = "" Then
        lblJunk(3).Caption = "&Hourly Reminders: Disabled"
    Else
        lblJunk(3).Caption = "&Hourly Reminders: Enabled"
    End If

    'Notification about missed tasks
        
    If Val(GetSetting("Upcoming", "Total", "Number", "0")) > 0 Then
        For nEvents = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To 0 Step -1
            If CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) >= Now Then
                Exit For
            ElseIf GetSetting("Upcoming", Trim(Str(nEvents)), "Clicked", "") = "" Then
                Load WINnotify
                WINnotify.lblContent.Caption = WINnotify.lblContent.Caption + GetSetting("Upcoming", Trim(Str(nEvents)), "Subject", "") + Chr(13)
                nEvents2 = nEvents2 + 1
                SaveSetting "Upcoming", Trim(Str(nEvents)), "Clicked", "1"
            End If
        Next
        If nEvents2 > 0 Then
            WINnotify.lblSubject.Caption = Str(nEvents2) + " Missed " + "task" + IIf(nEvents2 <> 1, "s", "")
            WINnotify.Tag = ""
            WINnotify.Show
            nEvents2 = 0
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdAddTask.BackColor = cButton_Normal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    BtnClose_Click
    Cancel = 1
End Sub

Private Sub lblAbout_Click()
    On Error Resume Next
    If Not JustHide Then
        WINabout.Show
    End If
    JustHide = False
End Sub

Private Sub lblAbout_DblClick()
    On Error Resume Next
    WINabout.Show
End Sub

Private Sub lblArchived_Click()
    On Error Resume Next
    Dim nEvents As Integer
    If TabSelected <> 1 Then
        txtContent.Visible = False
    End If
    If Not CalledExt And TabSelected = 1 Then
        Exit Sub
    End If
    Clear_All
    Load_All "Archived"
    For nEvents = WorksLoaded - 1 To 0 Step -1              'Adjustment of color upon work loading
        If CDate(lblTime(nEvents).Caption) < Now Then
            lblTime(nEvents).ForeColor = vbDarkRed
        Else
            lblTime(nEvents).ForeColor = vbBlack
        End If
    Next
    If WorksLoaded + IIf(txtContent.Visible, 3, 0) <= 5 Then
        lblScroll.Visible = False
    End If
    If CalledExt Then
        Exit Sub
    End If
    CalledExt = False
    lblDone.ToolTipText = "Restore"
    lblDone.Caption = "q"
    lblArchived.Height = 495
    lblArchived.Caption = ""
    lblArchived.Top = 960
    lblArchived.Width = 4455
    lblJunk(0).Caption = "Archived"
    lblJunk(0).Left = 600
    lblJunk(0).ZOrder 0
    lblArchived.ZOrder 0
    sTab(1).Left = 600
    sTab(1).Top = 960
    sTab(1).Width = 4455
    sTab(1).FillColor = cTab_Selected
    If TabSelected = 0 Then
        lblUpcoming.Top = 1080
        lblUpcoming.Height = 375
        lblUpcoming.Width = 2055
        lblUpcoming.Caption = "Upcoming"
        sTab(0).Top = 1080
        sTab(0).Width = 2055
        sTab(0).FillColor = cTab_Normal
        lblUpcoming.Left = 5055
        sTab(0).Left = 5055
    Else
        lblPending.Top = 1080
        lblPending.Height = 375
        lblPending.Width = 2055
        lblPending.Caption = "Pending"
        sTab(2).Top = 1080
        sTab(2).Width = 2055
        sTab(2).FillColor = cTab_Normal
        lblUpcoming.Left = 5055
        sTab(0).Left = 5055
        lblPending.Left = 7110
        sTab(2).Left = 7110
    End If
    TabSelected = 1
End Sub

Private Sub lblDelete_Click()
    On Error Resume Next
    If WorksLoaded > 0 Then
        Delete_Work TabSelected, WorkSelected
    End If
End Sub

Private Sub lblDone_Click()
    On Error Resume Next
    MarkDone TabSelected, WorkSelected
End Sub

Private Sub lblExpand_Click()
    On Error Resume Next
    Expand_Content WorkSelected
End Sub

Private Sub lblFeedback_Click()
    On Error Resume Next
    ShellExecute 0, "open", Chr(34) + "http://anico.in/apps/forums/topics/show/7589381-feedback" + Chr(34), 0, 0, 1
End Sub

Private Sub lblHourly_Click()
    On Error Resume Next
    If GetSetting("remindME", "Notify", "Latest Check", "") = "" Then
        SaveSetting "remindME", "Notify", "Latest Check", DateAdd("yyyy", -1, Now)
        lblJunk(3).Caption = "&Hourly Reminders: Enabled"
    Else
        DeleteSetting "remindME"
        lblJunk(3).Caption = "&Hourly Reminders: Disabled"
    End If
End Sub

Private Sub lblSubject_DblClick(Index As Integer)
    On Error Resume Next
    lblWork_DblClick Index
End Sub

Private Sub lblTime_Click(Index As Integer)
    On Error Resume Next
    lblWork_Click Index
End Sub

Public Sub Delete_Work(TabSel As Byte, WorkSel As Integer)
    On Error Resume Next
    Dim TempNum As Integer
    Dim WorkLoad As Integer
    Dim TempNumStr As String
    Select Case TabSel
        Case 0
            TempNumStr = "Upcoming"
        Case 1
            TempNumStr = "Archived"
        Case 2
            TempNumStr = "Pending"
    End Select
    TempNum = WorkSel
    WorkLoad = Val(GetSetting(TempNumStr, "Total", "Number", "0"))
    If WorkLoad < 1 Then
        Exit Sub
    End If
    DeleteSetting TempNumStr, Trim(Str(WorkSel))
    If WorkLoad > 1 Then
        For TempNum = WorkSel To WorkLoad - 2
            SaveSetting TempNumStr, Trim(Str(TempNum)), "Subject", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "Subject", "")
            lblSubject(TempNum).Caption = lblSubject(TempNum + 1).Caption
            SaveSetting TempNumStr, Trim(Str(TempNum)), "Content", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "Content", "")
            If TabSel <> 2 Then
                SaveSetting TempNumStr, Trim(Str(TempNum)), "Time", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "Time", "")
                lblTime(TempNum).Caption = lblTime(TempNum + 1).Caption
                SaveSetting TempNumStr, Trim(Str(TempNum)), "oTime", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "oTime", "")
                SaveSetting TempNumStr, Trim(Str(TempNum)), "Freq", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "Freq", "o")
                SaveSetting TempNumStr, Trim(Str(TempNum)), "Expiry", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "Expiry", Str(#12/31/2099#))
                SaveSetting TempNumStr, Trim(Str(TempNum)), "Clicked", GetSetting(TempNumStr, Trim(Str(TempNum + 1)), "Clicked", "")
            End If
        Next
    End If
    SaveSetting TempNumStr, GetSetting(TempNumStr, "Total", "Number", "0"), "TempNum", ""
    DeleteSetting TempNumStr, GetSetting(TempNumStr, "Total", "Number", "0")
    SaveSetting TempNumStr, "Total", "Number", Trim(Str(Val(GetSetting(TempNumStr, "Total", "Number", "0")) - 1))
    If TabSel = TabSelected Then
        CalledExt = True
        If WorkSel = ContentNo Then
            txtContent.Visible = False
        End If
        WorkSelectedTop = lblWork(WorkSelected).Top
        If WorkSelected Then
            WorkSelected = WorkSelected - 1
        End If
        Clear_All
        Load_All TempNumStr
        CalledExt = False
        AdjustWork
    End If
End Sub

Public Sub MarkDone(TabSel As Byte, WorkSel As Integer)
    On Error Resume Next
    Dim TempNum As Integer
    If TabSel <> 1 Then
        If TabSel = 2 Then
            SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Subject", GetSetting("Pending", Trim(Str(WorkSel)), "Subject", "")
            SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Content", GetSetting("Pending", Trim(Str(WorkSel)), "Content", "")
        Else
            If GetSetting("Upcoming", Trim(Str(WorkSel)), "Freq", "o") = "o" Or CDate(GetSetting("Upcoming", Trim(Str(WorkSel)), "Expiry", Str(#12/31/2099#))) < DateAdd(GetSetting("Upcoming", Trim(Str(WorkSel)), "Freq", "d"), 1, CDate(GetSetting("Upcoming", Trim(Str(WorkSel)), "Time", DateValue(Now)))) Then
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Subject", GetSetting("Upcoming", Trim(Str(WorkSel)), "Subject", "")
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Content", GetSetting("Upcoming", Trim(Str(WorkSel)), "Content", "")
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Time", GetSetting("Upcoming", Trim(Str(WorkSel)), "Time", "")
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "oTime", GetSetting("Upcoming", Trim(Str(WorkSel)), "Time", "")
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Freq", GetSetting("Upcoming", Trim(Str(WorkSel)), "Freq", "o")
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Expiry", GetSetting("Upcoming", Trim(Str(WorkSel)), "Expiry", Str(#12/31/2099#))
                SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Clicked", GetSetting("Upcoming", Trim(Str(WorkSel)), "Clicked", "")
            Else
                Dim TempSubject As String
                Dim TempContent As String
                Dim TempTime As String
                Dim TempFreq As String
                Dim TempExpiry As String
                TempSubject = GetSetting("Upcoming", Trim(Str(WorkSel)), "Subject", "")
                TempContent = GetSetting("Upcoming", Trim(Str(WorkSel)), "Content", "")
                TempTime = Str(DateAdd(GetSetting("Upcoming", Trim(Str(WorkSel)), "Freq"), 1, GetSetting("Upcoming", Trim(Str(WorkSel)), "oTime")))
                TempFreq = GetSetting("Upcoming", Trim(Str(WorkSel)), "Freq", "o")
                TempExpiry = GetSetting("Upcoming", Trim(Str(WorkSel)), "Expiry", Str(#12/31/2099#))
                Delete_Work TabSel, WorkSel
                AddTask TempSubject, TempContent, TempTime, TempFreq, TempExpiry
                Exit Sub
            End If
        End If
        SaveSetting "Archived", "Total", "Number", Trim(Str(Val(GetSetting("Archived", "Total", "Number", "0")) + 1))
        Delete_Work TabSel, WorkSel
    Else
        TempNum = WorkSel
        AddTask GetSetting("Archived", Trim(Str(WorkSel)), "Subject", ""), GetSetting("Archived", Trim(Str(WorkSel)), "Content", ""), GetSetting("Archived", Trim(Str(WorkSel)), "Time", ""), GetSetting("Archived", Trim(Str(WorkSel)), "Freq", "o"), GetSetting("Archived", Trim(Str(WorkSel)), "Expiry", Str(#12/31/2099#))
        CalledExt = True
        lblArchived_Click
        lblWork_Click TempNum
        Delete_Work TabSel, WorkSel
    End If
End Sub

Private Sub lblPending_Click()
    On Error Resume Next
    If TabSelected <> 2 Then
        txtContent.Visible = False
    End If
    If Not CalledExt And TabSelected = 2 Then
        Exit Sub
    End If
    Clear_All
    Load_All "Pending"
    If WorksLoaded + IIf(txtContent.Visible, 3, 0) <= 5 Then
        lblScroll.Visible = False
    End If
    If CalledExt Then
        Exit Sub
    End If
    CalledExt = False
    lblDone.ToolTipText = "Done"
    lblDone.Caption = "a"
    If TabSelected = 0 Then
        lblUpcoming.Top = 1080
        lblUpcoming.Height = 375
        lblUpcoming.Width = 2055
        lblUpcoming.Caption = "Upcoming"
        sTab(0).Top = 1080
        sTab(0).Width = 2055
        sTab(0).FillColor = cTab_Normal
    Else
        lblArchived.Top = 1080
        lblArchived.Height = 375
        lblArchived.Width = 2055
        lblArchived.Caption = "Archived"
        sTab(1).Top = 1080
        sTab(1).Width = 2055
        sTab(1).FillColor = cTab_Normal
        lblUpcoming.Left = 2655
        sTab(0).Left = 2655
    End If
    lblPending.Height = 495
    lblPending.Caption = ""
    lblPending.Top = 960
    lblPending.Width = 4455
    lblPending.Left = 4710
    lblJunk(0).Caption = "Pending"
    lblJunk(0).Left = 4710
    lblJunk(0).ZOrder 0
    lblPending.ZOrder 0
    sTab(2).Left = 4710
    sTab(2).Top = 960
    sTab(2).Width = 4455
    sTab(2).FillColor = cTab_Selected
    TabSelected = 2
End Sub

Private Sub lblScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblScroll.BackColor = cScroll_Pressed
    If Button = 1 Then
        Yinit = Y
        Mousedwn = True
    End If
End Sub

Private Sub lblScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblScroll.BackColor = cScroll_Hover
    If Button = 1 And Mousedwn Then
        lblScroll.Move lblScroll.Left, lblScroll.Top + Y - Yinit
        lblScroll.BackColor = cScroll_Pressed
        For TempNum = WorksLoaded - 1 To 0 Step -1
            lblWork(TempNum).Top = lblWork(TempNum).Top - ((Y - Yinit) * ((720 * (WorksLoaded - 5 + (IIf(txtContent.Visible = True, 3, 0)))) / 2880))
            If TabSelected <> 2 Then
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
            End If
            lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
        Next
        txtContent.Top = txtContent.Top - ((Y - Yinit) * ((720 * (WorksLoaded - 5 + (IIf(txtContent.Visible = True, 3, 0)))) / 2880))
        If lblDelete.Visible Then
            lblDelete.Top = lblWork(WorkSelected).Top + 180
        End If
        If lblDone.Visible Then
            lblDone.Top = lblWork(WorkSelected).Top + 150
        End If
        If lblExpand.Visible Then
            lblExpand.Top = lblWork(WorkSelected).Top + 150
        End If
    End If
End Sub

Private Sub lblScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblScroll.BackColor = cScroll_Normal
    If lblScroll.Top < 120 Then
        lblScroll.Top = 120
        lblWork(WorksLoaded - 1).Top = 120
        lblTime(WorksLoaded - 1).Top = 300
        lblSubject(WorksLoaded - 1).Top = 240
        For TempNum = WorksLoaded - 2 To 0 Step -1
            lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720 + IIf(TempNum + 1 = ContentNo And txtContent.Visible = True, 2160, 0)
            lblTime(TempNum).Top = lblWork(TempNum).Top + 180
            lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
        Next
    ElseIf lblScroll.Top > 3000 Then
        lblScroll.Top = 3000
        lblWork(0).Top = 3000 - IIf(ContentNo = 0 And txtContent.Visible = True, 2160, 0)
        lblTime(0).Top = lblWork(0).Top + 180
        lblSubject(0).Top = lblWork(0).Top + 120
        For TempNum = 1 To WorksLoaded - 1
            lblWork(TempNum).Top = lblWork(TempNum - 1).Top - 720 - IIf(TempNum = ContentNo And txtContent.Visible = True, 2160, 0)
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
            lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
        Next
    End If
    If lblDelete.Visible Then
        lblDelete.Top = lblWork(WorkSelected).Top + 180
    End If
    If lblDone.Visible Then
        lblDone.Top = lblWork(WorkSelected).Top + 150
    End If
    If lblExpand.Visible Then
        lblExpand.Top = lblWork(WorkSelected).Top + 150
    End If
    txtContent.Top = lblWork(ContentNo).Top + 720
End Sub

Private Sub lblSubject_Click(Index As Integer)
    On Error Resume Next
    lblWork_Click (Index)
End Sub

Private Sub lblTime_DblClick(Index As Integer)
    On Error Resume Next
    lblWork_DblClick Index
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Xinit = X
        Yinit = Y
        Mousedwn = True
    End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 And Mousedwn Then
        Move Left + X - Xinit, Top + Y - Yinit
    End If
End Sub

Private Sub lblUpcoming_Click()
    On Error Resume Next
    Dim nEvents As Integer
    If TabSelected Then
        txtContent.Visible = False
    End If
    If CalledExt = False And TabSelected = 0 Then
        Exit Sub
    End If
    Clear_All
    Load_All "Upcoming"
    For nEvents = WorksLoaded - 1 To 0 Step -1              'Adjustment of color upon work loading
        If CDate(lblTime(nEvents).Caption) < Now Then
            lblTime(nEvents).ForeColor = vbDarkRed
        Else
            lblTime(nEvents).ForeColor = vbBlack
        End If
    Next
    If WorksLoaded + IIf(txtContent.Visible, 3, 0) <= 5 Then
        lblScroll.Visible = False
    End If
    If CalledExt Then
        Exit Sub
    End If
    CalledExt = False
    lblDone.ToolTipText = "Done"
    lblDone.Caption = "a"
    AdjustWork
    If TabSelected = 1 Then
        lblArchived.Top = 1080
        lblArchived.Height = 375
        lblArchived.Width = 2055
        lblArchived.Caption = "Archived"
        sTab(1).Top = 1080
        sTab(1).Width = 2055
        sTab(1).FillColor = cTab_Normal
    Else
        lblPending.Top = 1080
        lblPending.Height = 375
        lblPending.Width = 2055
        lblPending.Caption = "Pending"
        sTab(2).Top = 1080
        sTab(2).Width = 2055
        sTab(2).FillColor = cTab_Normal
    End If
    lblUpcoming.Height = 495
    lblUpcoming.Caption = ""
    lblUpcoming.Top = 960
    lblUpcoming.Width = 4455
    lblUpcoming.Left = 2655
    lblJunk(0).Caption = "Upcoming"
    lblJunk(0).Left = 2655
    lblJunk(0).ZOrder 0
    lblUpcoming.ZOrder 0
    sTab(0).Left = 2655
    sTab(0).Top = 960
    sTab(0).Width = 4455
    sTab(0).FillColor = cTab_Selected
    lblPending.Left = 7110
    sTab(2).Left = 7110
    TabSelected = 0
End Sub
Private Sub lblWork_Click(Index As Integer)
    On Error Resume Next
    If Index = WorkSelected Then
        Exit Sub
    End If
    Dim HaveToExpand As Boolean
    On Error Resume Next
    lblWork(WorkSelected).BackColor = cButton_Normal
    lblWork(Index).BackColor = cButton_Hover
    If txtContent.Visible Then
        Expand_Content WorkSelected
        HaveToExpand = True
    End If
    WorkSelected = Index
    If lblWork(WorkSelected).Top < 120 Then
        lblWork(WorkSelected).Top = 120
        lblTime(WorkSelected).Top = 300
        lblSubject(WorkSelected).Top = 240
        For TempNum = WorkSelected + 1 To WorksLoaded - 1
            lblWork(TempNum).Top = lblWork(TempNum - 1).Top - 720
            lblTime(TempNum).Top = lblWork(TempNum).Top + 180
            lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
        Next
        If WorkSelected > 0 Then
            For TempNum = WorkSelected - 1 To 0 Step -1
                lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
        End If
        SetScroll
    End If
    If lblWork(WorkSelected).Top > 3000 Then
        lblWork(WorkSelected).Top = 3000
        lblTime(WorkSelected).Top = 3180
        lblSubject(WorkSelected).Top = 3120
        If WorkSelected < WorksLoaded - 1 Then
            For TempNum = WorkSelected + 1 To WorksLoaded - 1
                lblWork(TempNum).Top = lblWork(TempNum - 1).Top - 720
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
        End If
        For TempNum = WorkSelected - 1 To 0 Step -1
            lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720
            lblTime(TempNum).Top = lblWork(TempNum).Top + 180
            lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
        Next
        SetScroll
    End If
    lblDelete.Top = lblWork(WorkSelected).Top + 180
    lblDone.Top = lblWork(WorkSelected).Top + 150
    lblExpand.Top = lblWork(WorkSelected).Top + 150
    If HaveToExpand Then
        Expand_Content WorkSelected
    End If
End Sub
Private Sub Expand_Content(Index As Integer)
    On Error Resume Next
    Dim i, temp As Integer
    temp = Index
    If Not txtContent.Visible Then
        If lblWork(Index).Top + 720 > 1560 Then
            txtContent.Top = 1560
            lblWork(Index).Top = txtContent.Top - 720
            lblTime(Index).Top = lblWork(Index).Top + 180
            lblSubject(Index).Top = lblWork(Index).Top + 120
            lblDelete.Top = lblWork(WorkSelected).Top + 180
            lblDone.Top = lblWork(WorkSelected).Top + 150
            lblExpand.Top = lblWork(WorkSelected).Top + 150
            For i = Index + 1 To WorksLoaded - 1
                lblWork(i).Top = lblWork(i - 1).Top - 720
                lblTime(i).Top = lblWork(i).Top + 180
                lblSubject(i).Top = lblWork(i).Top + 120
            Next
        Else
            txtContent.Top = lblWork(Index).Top + 720
        End If
        lblWork(Index - 1).Top = txtContent.Top + 2160
        lblTime(Index - 1).Top = lblWork(Index - 1).Top + 180
        lblSubject(Index - 1).Top = lblWork(Index - 1).Top + 120
        For i = Index - 2 To 0 Step -1
            lblWork(i).Top = lblWork(i + 1).Top + 720
            lblTime(i).Top = lblWork(i).Top + 180
            lblSubject(i).Top = lblWork(i).Top + 120
        Next
        txtContent.Visible = True
        lblExpand.Caption = "6"
        lblExpand.ToolTipText = "Collapse"
        txtContent.SetFocus
        ContentNo = Index
        LoadContent
    Else
        txtContent.Visible = False
        lblExpand.Caption = "4"
        lblExpand.ToolTipText = "Expand"
        If ContentNo = Index Then
            For i = Index - 1 To 0 Step -1
                lblWork(i).Top = lblWork(i).Top - 2160
                lblTime(i).Top = lblWork(i).Top + 180
                lblSubject(i).Top = lblWork(i).Top + 120
            Next
        End If
    End If
    AdjustWork
    If WorksLoaded + IIf(txtContent.Visible, 3, 0) > 5 Then
        lblScroll.Visible = True
    Else
        lblScroll.Visible = False
    End If
    AdjustWork
    SetScroll
End Sub
Private Sub AdjustWork()
    On Error Resume Next
    If lblWork(0).Top < 3000 Then
        If WorksLoaded + (IIf(txtContent.Visible, 3, 0)) > 4 Then
            lblScroll.Top = 3000
            lblWork(0).Top = 3000 - IIf(ContentNo = 0 And txtContent.Visible = True, 2160, 0)
            lblTime(0).Top = lblWork(0).Top + 180
            lblSubject(0).Top = lblWork(0).Top + 120
            For TempNum = 1 To WorksLoaded - 1
                lblWork(TempNum).Top = lblWork(TempNum - 1).Top - 720 - IIf(TempNum = ContentNo And txtContent.Visible = True, 2160, 0)
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
            SetScroll
    
        ElseIf lblWork(WorksLoaded - 1).Top < 120 Then
            lblScroll.Top = 120
            lblWork(WorksLoaded - 1).Top = 120
            lblTime(WorksLoaded - 1).Top = 300
            lblSubject(WorksLoaded - 1).Top = 240
            For TempNum = WorksLoaded - 2 To 0 Step -1
                lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720 + IIf(TempNum + 1 = ContentNo And txtContent.Visible = True, 2160, 0)
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
        End If
    End If
    If lblWork(WorksLoaded - 1).Top > 120 Then
            lblScroll.Top = 120
            lblWork(WorksLoaded - 1).Top = 120
            lblTime(WorksLoaded - 1).Top = 300
            lblSubject(WorksLoaded - 1).Top = 240
            For TempNum = WorksLoaded - 2 To 0 Step -1
                lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720 + IIf(TempNum + 1 = ContentNo And txtContent.Visible = True, 2160, 0)
                lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
    End If
    If lblDelete.Visible Then
            lblDelete.Top = lblWork(WorkSelected).Top + 180
    End If
    If lblDone.Visible Then
        lblDone.Top = lblWork(WorkSelected).Top + 150
    End If
    If lblExpand.Visible Then
        lblExpand.Top = lblWork(WorkSelected).Top + 150
    End If
End Sub
Private Sub SetScroll()
    On Error Resume Next
    lblScroll.Top = 120 + (120 - lblWork(WorksLoaded - 1).Top) / ((720 * (WorksLoaded - 5 + (IIf(txtContent.Visible = True, 3, 0)))) / 2880)
End Sub
Private Sub lblWork_DblClick(Index As Integer)
    On Error Resume Next
    If TabSelected <> 1 Then
        Edit_Work
    End If
End Sub

Public Sub AddTask(s_Subject As String, s_Content As String, Optional s_Date As String = "", Optional s_Freq As String = "o", Optional s_Expiry As String = "", Optional s_oTime As String = "")
    On Error Resume Next
    If s_Subject = "" Then
        Exit Sub
    End If
    If s_Expiry = "" Then
        s_Expiry = Str(#12/31/2099#)
    End If
    If s_oTime = "" Then
        s_oTime = s_Date
    End If
    If s_Date = "" Then                     'The case for pending tasks
        SaveSetting "Pending", GetSetting("Pending", "Total", "Number", "0"), "Subject", s_Subject
        SaveSetting "Pending", GetSetting("Pending", "Total", "Number", "0"), "Content", s_Content
        SaveSetting "Pending", "Total", "Number", Trim(Str(Val(GetSetting("Pending", "Total", "Number", "0")) + 1))
        WorkSelected = Val(GetSetting("Pending", "Total", "Number", "0")) - 1
        Clear_All
        CalledExt = True
        lblPending_Click
    Else                                    'The case for upcoming tasks
        TempNum2 = 0
        If Val(GetSetting("Upcoming", "Total", "Number", "0")) > 0 Then
            For TempNum2 = 0 To Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1
                If CDate(s_Date) > CDate(GetSetting("Upcoming", Trim(Str(TempNum2)), "Time")) Then
                    Exit For
                End If
            Next
            For TempNum3 = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To TempNum2 Step -1
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "Subject", GetSetting("Upcoming", Trim(Str(TempNum3)), "Subject", "")
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "Content", GetSetting("Upcoming", Trim(Str(TempNum3)), "Content", "")
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "Time", GetSetting("Upcoming", Trim(Str(TempNum3)), "Time", "")
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "oTime", GetSetting("Upcoming", Trim(Str(TempNum3)), "oTime", "")
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "Freq", GetSetting("Upcoming", Trim(Str(TempNum3)), "Freq", "o")
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "Expiry", GetSetting("Upcoming", Trim(Str(TempNum3)), "Expiry", Str(#12/31/2099#))
                SaveSetting "Upcoming", Trim(Str(TempNum3 + 1)), "Clicked", GetSetting("Upcoming", Trim(Str(TempNum3)), "Clicked", "")
            Next
        End If
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "Subject", s_Subject
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "Content", s_Content
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "Time", s_Date
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "oTime", s_oTime
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "Freq", s_Freq
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "Expiry", s_Expiry
        SaveSetting "Upcoming", Trim(Str(TempNum2)), "Clicked", ""
        SaveSetting "Upcoming", "Total", "Number", Trim(Str(Val(GetSetting("Upcoming", "Total", "Number", "0")) + 1))
        Dim X As Integer
        X = TempNum2
        Clear_All
        CalledExt = True
        lblUpcoming_Click
        lblWork_Click X
    End If
    AdjustWork
End Sub

Public Sub Clear_All()
    On Error Resume Next
    For TempNum2 = lblWork.LBound + 1 To lblWork.UBound
        Unload lblWork(TempNum2)
    Next
    For TempNum2 = lblTime.LBound + 1 To lblTime.UBound
        Unload lblTime(TempNum2)
    Next
    For TempNum2 = lblSubject.LBound + 1 To lblSubject.UBound
        Unload lblSubject(TempNum2)
    Next
    lblDelete.Visible = False
    lblDone.Visible = False
    lblExpand.Visible = False
    lblWork(0).Visible = False
    lblSubject(0).Visible = False
    lblTime(0).Visible = False
    lblScroll.Visible = False
    WorksLoaded = 0
End Sub

Public Sub Load_All(sName As String)
    On Error Resume Next
    txtContent.Visible = False
    lblWork(0).BackColor = cButton_Normal
    If Val(GetSetting(sName, "Total", "Number", "0")) = 0 Then
        Exit Sub
    End If
    For TempNum3 = 1 To Val(GetSetting(sName, "Total", "Number", "0")) - 1
        Load lblWork(TempNum3)
        Load lblTime(TempNum3)
        Load lblSubject(TempNum3)
    Next
    lblSubject(TempNum3 - 1).Caption = GetSetting(sName, Trim(Str(Val(GetSetting(sName, "Total", "Number", "0")) - 1)), "Subject", "")
    lblTime(TempNum3 - 1).Caption = GetSetting(sName, Trim(Str(Val(GetSetting(sName, "Total", "Number", "0")) - 1)), "Time", "")
    lblTime(TempNum3 - 1).Left = 240
    lblTime(TempNum3 - 1).ZOrder 0
    lblWork(TempNum3 - 1).Left = 120
    lblSubject(TempNum3 - 1).Left = 3120
    lblSubject(TempNum3 - 1).ZOrder 0
    For TempNum3 = TempNum3 - 2 To 0 Step -1
        lblSubject(TempNum3).Caption = GetSetting(sName, Trim(Str(TempNum3)), "Subject", "")
        lblTime(TempNum3).Caption = GetSetting(sName, Trim(Str(TempNum3)), "Time", "")
        lblTime(TempNum3).ZOrder 0
        lblSubject(TempNum3).ZOrder 0
    Next
    WorksLoaded = Val(GetSetting(sName, "Total", "Number", "0"))
    If WorksLoaded > 5 Then
        lblScroll.Visible = True
    Else
        lblScroll.Visible = False
    End If
    If Not CalledExt Then
        WorkSelected = WorksLoaded - 1
        WorkSelectedTop = 120
    End If
    lblWork(WorkSelected).BackColor = cButton_Hover
    lblWork(WorkSelected).Top = WorkSelectedTop
    lblSubject(WorkSelected).Top = lblWork(WorkSelected).Top + 120
    lblTime(WorkSelected).Top = lblWork(WorkSelected).Top + 180
    If CDate(lblTime(WorkSelected).Caption) < Now Then
        lblTime(WorkSelected).ForeColor = vbDarkRed
    Else
        lblTime(WorkSelected).ForeColor = vbBlack
    End If
    For TempNum3 = WorkSelected + 1 To WorksLoaded - 1
        lblWork(TempNum3).Top = lblWork(TempNum3 - 1).Top - 720
        lblSubject(TempNum3).Top = lblWork(TempNum3).Top + 120
        lblTime(TempNum3).Top = lblWork(TempNum3).Top + 180
        If CDate(lblTime(TempNum3).Caption) < Now Then
            lblTime(TempNum3).ForeColor = vbDarkRed
        Else
            lblTime(TempNum3).ForeColor = vbBlack
        End If
    Next
    For TempNum3 = WorkSelected - 1 To 0 Step -1
        lblWork(TempNum3).Top = lblWork(TempNum3 + 1).Top + 720
        lblSubject(TempNum3).Top = lblWork(TempNum3).Top + 120
        lblTime(TempNum3).Top = lblWork(TempNum3).Top + 180
        If CDate(lblTime(TempNum3).Caption) < Now Then
            lblTime(TempNum3).ForeColor = vbDarkRed
        Else
            lblTime(TempNum3).ForeColor = vbBlack
        End If
    Next
    For TempNum3 = WorksLoaded - 1 To 0 Step -1
        lblWork(TempNum3).Visible = True
        lblSubject(TempNum3).Visible = True
        lblTime(TempNum3).Visible = True
    Next
    lblDelete.Top = lblWork(WorkSelected).Top + 180
    lblDone.Top = lblDelete.Top - 30
    lblExpand.Top = lblDone.Top
    lblDelete.Visible = True
    lblDone.Visible = True
    lblExpand.Visible = True
    lblExpand.Caption = "4"
    lblExpand.ToolTipText = "Expand"
    CalledExt = False
    SetScroll
End Sub

Private Sub Notify_Timer()
    On Error Resume Next
    JustHide = False
    
    'Check for Dead events, make them dark red
    
    If TabSelected <> 2 Then
        For nEvents = WorksLoaded - 1 To 0 Step -1
            If CDate(lblTime(nEvents).Caption) < Now Then
                lblTime(nEvents).ForeColor = vbDarkRed
            Else
                lblTime(nEvents).ForeColor = vbBlack
            End If
        Next
    End If
    
    'Hourly Notifications of upcoming tasks
    
    If Not WINnotifyLoaded Then
        If GetSetting("remindME", "Notify", "Latest Check", "") <> "" Then
            If Not IsDate(GetSetting("remindME", "Notify", "Latest Check", "")) Then
                SaveSetting "remindME", "Notify", "Latest Check", DateAdd("yyyy", -1, Now)
            End If
            If Val(Now - CDate(GetSetting("remindME", "Notify", "Latest Check", ""))) > OneHr Then
                If Val(GetSetting("Upcoming", "Total", "Number", "0")) > 0 Then
                    For nEvents = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To 0 Step -1
                        If Val(CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) - Now) > OneHr Then
                            Exit For
                        ElseIf CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) > Now Then
                            Load WINnotify
                            WINnotify.lblContent.Caption = WINnotify.lblContent.Caption + GetSetting("Upcoming", Trim(Str(nEvents)), "Subject", "") + Chr(13)
                            nEvents2 = nEvents2 + 1
                        End If
                    Next
                    If nEvents2 > 0 Then
                        WINnotify.lblSubject.Caption = Trim(Str(nEvents2)) + " upcoming " + "task" + IIf(nEvents2 <> 1, "s", "") + " in next one hour"
                        WINnotify.Tag = ""
                        WINnotify.Show
                        nEvents2 = 0
                    End If
                End If
                SaveSetting "remindME", "Notify", "Latest Check", Str(Now)
            End If
        End If
    End If
    
    'Notification of a specific task
    
    If Not WINnotifyLoaded Then
        For nEvents = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To 0 Step -1
            If CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) < Now And GetSetting("Upcoming", Trim(Str(nEvents)), "Clicked", "") = "" Then
                Load WINnotify
                WINnotify.lblContent.Caption = GetSetting("Upcoming", Trim(Str(nEvents)), "Content", "")
                WINnotify.lblSubject.Caption = GetSetting("Upcoming", Trim(Str(nEvents)), "Subject", "")
                WINnotify.lblContent.Top = WINnotify.lblSubject.Top + WINnotify.lblSubject.Height + 105
                If WINnotify.lblContent.Height + WINnotify.lblContent.Top > 2175 Then
                    WINnotify.Height = WINnotify.lblContent.Height + WINnotify.lblContent.Top + 690
                End If
                For nEvents2 = 0 To 5
                    WINnotify.lblPostpone(nEvents2).Top = WINnotify.Height - 465
                Next
                nEvents2 = 0
                WINnotify.Tag = Trim(Str(nEvents))
                WINnotify.Show
            End If
        Next
    End If
End Sub

Private Sub LoadContent()
    On Error Resume Next
    Dim FreqStr As String
    HeaderLen = 0
    Select Case TabSelected
        Case 0
            Select Case GetSetting("Upcoming", Trim(Str(ContentNo)), "Freq", "o")
                Case "o"
                    FreqStr = "once"
                Case "d"
                    FreqStr = "day"
                Case "ww"
                    FreqStr = "week"
                Case "m"
                    FreqStr = "month"
                Case "yyyy"
                    FreqStr = "year"
            End Select
            If FreqStr <> "once" Then
                txtContent.Text = "## This task is scheduled to occur every " + FreqStr + " expiring on " + GetSetting("Upcoming", Trim(Str(ContentNo)), "Expiry", Str(#12/31/2099#)) + " ##" + vbCrLf + vbCrLf
                HeaderLen = Len(txtContent.Text)
                txtContent.Text = txtContent.Text + GetSetting("Upcoming", Trim(Str(ContentNo)), "Content", "")
            Else
                txtContent.Text = GetSetting("Upcoming", Trim(Str(ContentNo)), "Content", "")
            End If
        Case 1
            Select Case GetSetting("Archived", Trim(Str(ContentNo)), "Freq", "o")
                Case "o"
                    FreqStr = "once"
                Case "d"
                    FreqStr = "day"
                Case "ww"
                    FreqStr = "week"
                Case "m"
                    FreqStr = "month"
                Case "yyyy"
                    FreqStr = "year"
            End Select
            If FreqStr <> "once" Then
                If CDate(GetSetting("Archived", Trim(Str(ContentNo)), "Expiry", Str(#12/31/2099#))) > Now Then
                    txtContent.Text = "## This task was scheduled to occur every " + FreqStr + " expiring on " + GetSetting("Archived", Trim(Str(ContentNo)), "Expiry", Str(#12/31/2099#)) + " ##" + vbCrLf + vbCrLf
                    HeaderLen = Len(txtContent.Text)
                    txtContent.Text = txtContent.Text + GetSetting("Archived", Trim(Str(ContentNo)), "Content", "")
                Else
                    txtContent.Text = "## This task was scheduled to occur every " + FreqStr + " expired on " + GetSetting("Archived", Trim(Str(ContentNo)), "Expiry", Str(#12/31/2099#)) + " ##" + vbCrLf + vbCrLf
                    HeaderLen = Len(txtContent.Text)
                    txtContent.Text = txtContent.Text + GetSetting("Archived", Trim(Str(ContentNo)), "Content", "")
                End If
            Else
                txtContent.Text = GetSetting("Archived", Trim(Str(ContentNo)), "Content", "")
            End If
        Case 2
            txtContent.Text = GetSetting("Pending", Trim(Str(ContentNo)), "Content", "")
    End Select
End Sub

Private Sub UpcomingFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If WorksLoaded > 0 Then
        lblScroll.BackColor = cScroll_Normal
        If lblWork(WorksLoaded - 1).Top > 120 Then
            lblScroll.Top = 120
            lblWork(WorksLoaded - 1).Top = 120
            If TabSelected <> 2 Then
                lblTime(WorksLoaded - 1).Top = 300
            End If
            lblSubject(WorksLoaded - 1).Top = 240
            For TempNum = WorksLoaded - 2 To 0 Step -1
                lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720
                If TabSelected <> 2 Then
                    lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                End If
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
        ElseIf lblWork(0).Top < 3000 And WorksLoaded > 5 Then
            lblScroll.Top = 3000
            lblWork(0).Top = 3000
            lblTime(0).Top = 3180
            lblSubject(0).Top = 3120
            For TempNum = WorksLoaded - 2 To 0 Step -1
                lblWork(TempNum).Top = lblWork(TempNum + 1).Top + 720
                If TabSelected <> 2 Then
                    lblTime(TempNum).Top = lblWork(TempNum).Top + 180
                End If
                lblSubject(TempNum).Top = lblWork(TempNum).Top + 120
            Next
        End If
        If lblDelete.Visible Then
            lblDelete.Top = lblWork(WorkSelected).Top + 180
        End If
        If lblDone.Visible Then
            lblDone.Top = lblWork(WorkSelected).Top + 150
        End If
        If lblExpand.Visible Then
            lblExpand.Top = lblWork(WorkSelected).Top + 150
        End If
    End If
End Sub
Private Sub Edit_Work()
    On Error Resume Next
    Load WINtask
    WINtask.Tag = "Edit"
    WINtask.Left = Left + (Width - WINtask.Width) / 2
    WINtask.Top = Top + (Height - WINtask.Height) / 2
    WINtask.txtSubject.Text = lblSubject(WorkSelected).Caption
    If TabSelected <> 2 Then
        WINtask.txtDeadline.Text = DateValue(lblTime(WorkSelected).Caption)
        WINtask.txtTime.Text = TimeValue(lblTime(WorkSelected).Caption)
        If GetSetting("Upcoming", Trim(Str(WorkSelected)), "Freq", "o") <> "o" Then
            WINtask.FreqO.BackColor = cButton_Normal
            WINtask.FreqO.ForeColor = vbGray
            WINtask.cmdOK.Tag = GetSetting("Upcoming", Trim(Str(WorkSelected)), "Freq", "o")
            WINtask.ShowExpiry
            WINtask.lblJunk(2).Caption = "From Date*"
            WINtask.txtExpiry.Text = GetSetting("Upcoming", Trim(Str(WorkSelected)), "Expiry", Str(#12/31/2099#))
        End If
        Select Case GetSetting("Upcoming", Trim(Str(WorkSelected)), "Freq", "o")
            Case "d"
                WINtask.FreqD.BackColor = cButton_Hover
                WINtask.FreqD.ForeColor = vbDarkGray
            Case "ww"
                WINtask.FreqW.BackColor = cButton_Hover
                WINtask.FreqW.ForeColor = vbDarkGray
            Case "m"
                WINtask.FreqM.BackColor = cButton_Hover
                WINtask.FreqM.ForeColor = vbDarkGray
            Case "yyyy"
                WINtask.FreqY.BackColor = cButton_Hover
                WINtask.FreqY.ForeColor = vbDarkGray
        End Select
    End If
    Select Case TabSelected
        Case 0
            WINtask.txtContent.Text = GetSetting("Upcoming", Trim(Str(WorkSelected)), "Content", "")
        Case 1
            WINtask.txtContent.Text = GetSetting("Archived", Trim(Str(WorkSelected)), "Content", "")
        Case 2
            WINtask.txtContent.Text = GetSetting("Pending", Trim(Str(WorkSelected)), "Content", "")
    End Select
    WINtask.Show vbModal
    If WINtask.Tag = "Editted" Then
        Delete_Work TabSelected, WorkSelected
        AddTask WINtask.txtSubject.Text, WINtask.txtContent.Text, WINtask.txtDeadline.Text, WINtask.cmdOK.Tag, WINtask.txtExpiry.Text
    End If
    Unload WINtask
End Sub
