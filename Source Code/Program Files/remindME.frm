VERSION 5.00
Begin VB.Form WINremindME 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "remindME"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   Icon            =   "remindME.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "remindME.frx":89EA
   ScaleHeight     =   6540
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Notify 
      Interval        =   1000
      Left            =   6840
      Top             =   5520
   End
   Begin VB.Frame ContentFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00125D82&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   8550
      Begin VB.Label lblDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblDelete 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7920
         TabIndex        =   12
         Top             =   300
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblSubject 
         AutoSize        =   -1  'True
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
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   60
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
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblWork 
         BackColor       =   &H001778A9&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.Label lblScroll 
         BackColor       =   &H001D719A&
         Height          =   615
         Left            =   8280
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Label BtnClose 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   9330
      TabIndex        =   11
      Top             =   180
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
      TabIndex        =   10
      Top             =   1080
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
      TabIndex        =   9
      Top             =   1080
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
      TabIndex        =   7
      Top             =   960
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
   Begin VB.Label cmdOK 
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   0
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
      TabIndex        =   8
      Top             =   1050
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
'Declarations for Borderless form
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_DLGFRAME = &H400000
Dim Xinit As Long
Dim Yinit As Long
Dim Mousedwn As Boolean
Dim WorksLoaded As Integer
Dim temp As Integer
Dim temp2 As Integer
Dim temp3 As Integer
Public TabSelected As Byte
Dim Vchanged As Boolean
Public WorkSelected As Integer
Dim CalledExt As Boolean
Const OneHr As Double = 4.16666666666667E-02
Const OneSec As Double = 6.94444444444444E-04
Private Sub BtnClose_Click()
    On Error Resume Next
    WindowState = 1
End Sub
Private Sub cmdOK_Click()
    On Error Resume Next
    cmdOK.BackColor = &H1778A9
    Load WINtask
    WINtask.Tag = "Add"
    WINtask.Left = Left + (Width - WINtask.Width) / 2
    WINtask.Top = Top + (Height - WINtask.Height) / 2
    WINtask.Show vbModal
    Call AddTask(WINtask.txtSubject.Text, WINtask.txtContent.Text, WINtask.txtDeadline.Text)
    Unload WINtask
End Sub
Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdOK.BackColor = RGB(225, 165, 38)
End Sub
Private Sub ContentFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If WorksLoaded > 0 Then
        lblScroll.BackColor = RGB(154, 113, 29)
        If lblWork(WorksLoaded - 1).Top > 120 Then
            lblScroll.Top = 120
            lblWork(WorksLoaded - 1).Top = 120
            If TabSelected <> 2 Then
                lblTime(WorksLoaded - 1).Top = 300
            End If
            lblSubject(WorksLoaded - 1).Top = 240
            For temp = WorksLoaded - 2 To 0 Step -1
                lblWork(temp).Top = lblWork(temp + 1).Top + 720
                If TabSelected <> 2 Then
                    lblTime(temp).Top = lblWork(temp).Top + 180
                End If
                lblSubject(temp).Top = lblWork(temp).Top + 120
            Next
        ElseIf lblWork(0).Top < 3000 And WorksLoaded > 5 Then
            lblScroll.Top = 3000
            lblWork(0).Top = 3000
            lblTime(0).Top = 3180
            lblSubject(0).Top = 3120
            For temp = WorksLoaded - 2 To 0 Step -1
                lblWork(temp).Top = lblWork(temp + 1).Top + 720
                If TabSelected <> 2 Then
                    lblTime(temp).Top = lblWork(temp).Top + 180
                End If
                lblSubject(temp).Top = lblWork(temp).Top + 120
            Next
        End If
        If lblDelete.Visible = True Then
            lblDelete.Top = lblWork(WorkSelected).Top + 180
        End If
        If lblDone.Visible = True Then
            lblDone.Top = lblWork(WorkSelected).Top + 180
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyTab
            Select Case TabSelected
                Case 0
                    If Shift = 0 Then
                        Call lblPending_Click
                    Else
                        Call lblArchived_Click
                    End If
                Case 1
                    If Shift = 0 Then
                        Call lblUpcoming_Click
                    Else
                        Call lblPending_Click
                    End If
                Case 2
                    If Shift = 0 Then
                        Call lblArchived_Click
                    Else
                        Call lblUpcoming_Click
                    End If
            End Select
        Case vbKeyA
            Call cmdOK_Click
        Case vbKeyDown
            If WorkSelected > 0 Then
                lblWork(WorkSelected).BackColor = &H1778A9
                lblWork(WorkSelected - 1).BackColor = RGB(225, 165, 38)
                WorkSelected = WorkSelected - 1
                If lblWork(WorkSelected).Top > 3000 Then
                    lblWork(WorkSelected).Top = 3000
                    If TabSelected <> 2 Then
                        lblTime(WorkSelected).Top = 3180
                    End If
                    lblSubject(WorkSelected).Top = 3120
                    For temp = WorkSelected - 1 To 0 Step -1
                        lblWork(temp).Top = lblWork(temp + 1).Top + 720
                        If TabSelected <> 2 Then
                            lblTime(temp).Top = lblWork(temp).Top + 180
                        End If
                        lblSubject(temp).Top = lblWork(temp).Top + 120
                    Next
                    For temp = WorkSelected + 1 To WorksLoaded - 1
                        lblWork(temp).Top = lblWork(temp - 1).Top - 720
                        If TabSelected <> 2 Then
                            lblTime(temp).Top = lblWork(temp).Top + 180
                        End If
                        lblSubject(temp).Top = lblWork(temp).Top + 120
                    Next
                    lblScroll.Top = 3000 - (lblWork(0).Top - lblWork(WorkSelected).Top) * 2880 / (720 * (WorksLoaded - 5))
                End If
            End If
        Case vbKeyUp
            If WorkSelected < WorksLoaded - 1 Then
                lblWork(WorkSelected).BackColor = &H1778A9
                lblWork(WorkSelected + 1).BackColor = RGB(225, 165, 38)
                WorkSelected = WorkSelected + 1
                If lblWork(WorkSelected).Top < 120 Then
                    lblWork(WorkSelected).Top = 120
                    If TabSelected <> 2 Then
                        lblTime(WorkSelected).Top = 300
                    End If
                    lblSubject(WorkSelected).Top = 240
                    If WorkSelected < WorksLoaded - 1 Then
                        For temp = WorkSelected + 1 To WorksLoaded - 1
                            lblWork(temp).Top = lblWork(temp - 1).Top - 720
                            If TabSelected <> 2 Then
                                lblTime(temp).Top = lblWork(temp).Top + 180
                            End If
                            lblSubject(temp).Top = lblWork(temp).Top + 120
                        Next
                    End If
                    For temp = WorkSelected - 1 To 0 Step -1
                        lblWork(temp).Top = lblWork(temp + 1).Top + 720
                        If TabSelected <> 2 Then
                            lblTime(temp).Top = lblWork(temp).Top + 180
                        End If
                        lblSubject(temp).Top = lblWork(temp).Top + 120
                    Next
                    lblScroll.Top = 120 + (lblWork(WorkSelected).Top - lblWork(WorksLoaded - 1).Top) * 2880 / (720 * (WorksLoaded - 5))
                End If
            End If
        Case vbKeyHome
            lblWork(WorkSelected).BackColor = &H1778A9
            lblWork(WorksLoaded - 1).BackColor = RGB(225, 165, 38)
            WorkSelected = WorksLoaded - 1
            lblScroll.Top = 120
            If lblWork(WorksLoaded - 1).Top < 120 Then
                lblWork(WorksLoaded - 1).Top = 120
                If TabSelected <> 2 Then
                    lblTime(WorksLoaded - 1).Top = 300
                End If
                lblSubject(WorksLoaded - 1).Top = 240
                For temp = WorksLoaded - 2 To 0 Step -1
                    lblWork(temp).Top = lblWork(temp + 1).Top + 720
                    If TabSelected <> 2 Then
                        lblTime(temp).Top = lblWork(temp).Top + 180
                    End If
                    lblSubject(temp).Top = lblWork(temp).Top + 120
                Next
            End If
        Case vbKeyEnd
            lblWork(WorkSelected).BackColor = &H1778A9
            lblWork(0).BackColor = RGB(225, 165, 38)
            WorkSelected = 0
            lblScroll.Top = 3000
            If lblWork(0).Top > 3000 Then
                lblWork(0).Top = 3000
                lblTime(0).Top = 3180
                lblSubject(0).Top = 3120
                For temp = 1 To WorksLoaded - 1
                    lblWork(temp).Top = lblWork(temp - 1).Top - 720
                    If TabSelected <> 2 Then
                        lblTime(temp).Top = lblWork(temp).Top + 180
                    End If
                    lblSubject(temp).Top = lblWork(temp).Top + 120
                Next
            End If
        Case vbKeyPageUp
            If WorkSelected < WorksLoaded - 1 Then
                lblWork(WorkSelected).BackColor = &H1778A9
                WorkSelected = IIf(WorkSelected + 5 < WorksLoaded, WorkSelected + 5, WorksLoaded - 1)
                lblWork(WorkSelected).BackColor = RGB(225, 165, 38)
                If lblWork(WorkSelected).Top < 120 Then
                    lblWork(WorkSelected).Top = 120
                    If TabSelected <> 2 Then
                        lblTime(WorkSelected).Top = 300
                    End If
                    lblSubject(WorkSelected).Top = 240
                    If WorkSelected < WorksLoaded - 1 Then
                        For temp = WorkSelected + 1 To WorksLoaded - 1
                            lblWork(temp).Top = lblWork(temp - 1).Top - 720
                            If TabSelected <> 2 Then
                                lblTime(temp).Top = lblWork(temp).Top + 180
                            End If
                            lblSubject(temp).Top = lblWork(temp).Top + 120
                        Next
                    End If
                    For temp = WorkSelected - 1 To 0 Step -1
                        lblWork(temp).Top = lblWork(temp + 1).Top + 720
                        If TabSelected <> 2 Then
                            lblTime(temp).Top = lblWork(temp).Top + 180
                        End If
                        lblSubject(temp).Top = lblWork(temp).Top + 120
                    Next
                    lblScroll.Top = 120 + (lblWork(WorkSelected).Top - lblWork(WorksLoaded - 1).Top) * 2880 / (720 * (WorksLoaded - 5))
                End If
            End If
        Case vbKeyPageDown
            If WorkSelected > 0 Then
                lblWork(WorkSelected).BackColor = &H1778A9
                WorkSelected = IIf(WorkSelected - 5 >= 0, WorkSelected - 5, 0)
                lblWork(WorkSelected).BackColor = RGB(225, 165, 38)
                If lblWork(WorkSelected).Top > 3000 Then
                    lblWork(WorkSelected).Top = 3000
                    If TabSelected <> 2 Then
                        lblTime(WorkSelected).Top = 3180
                    End If
                    lblSubject(WorkSelected).Top = 3120
                    For temp = WorkSelected + 1 To WorksLoaded - 1
                        lblWork(temp).Top = lblWork(temp - 1).Top - 720
                        If TabSelected <> 2 Then
                            lblTime(temp).Top = lblWork(temp).Top + 180
                        End If
                        lblSubject(temp).Top = lblWork(temp).Top + 120
                    Next
                    If WorkSelected > 0 Then
                        For temp = WorkSelected - 1 To 0 Step -1
                            lblWork(temp).Top = lblWork(temp + 1).Top + 720
                            If TabSelected <> 2 Then
                                lblTime(temp).Top = lblWork(temp).Top + 180
                            End If
                            lblSubject(temp).Top = lblWork(temp).Top + 120
                        Next
                    End If
                    lblScroll.Top = 3000 - (lblWork(0).Top - lblWork(WorkSelected).Top) * 2880 / (720 * (WorksLoaded - 5))
                End If
            End If
        Case vbKeyDelete
            Call lblDelete_Click
        Case vbKeyReturn
            If WorksLoaded > 0 Then
                Call lblWork_DblClick(WorkSelected)
            End If
    End Select
    If WorksLoaded > 0 Then
        lblDelete.Top = lblWork(WorkSelected).Top + 180
        lblDone.Top = lblWork(WorkSelected).Top + 180
    End If
End Sub
Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance = True Then
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
    CalledExt = False
    Load_All ("Upcoming")
    Call Notify_Timer
    DoEvents
    WindowState = 1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Call ContentFrame_MouseMove(0, 0, 0, 0)
    cmdOK.BackColor = RGB(169, 120, 23)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call BtnClose_Click
    Cancel = 1
End Sub
Private Sub lblArchived_Click()
    On Error Resume Next
    If CalledExt = False And TabSelected = 1 Then
        Exit Sub
    End If
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
    sTab(1).FillColor = &H125D82
    If TabSelected = 0 Then
        lblUpcoming.Top = 1080
        lblUpcoming.Height = 375
        lblUpcoming.Width = 2055
        lblUpcoming.Caption = "Upcoming"
        sTab(0).Top = 1080
        sTab(0).Width = 2055
        sTab(0).FillColor = &H95075
        lblUpcoming.Left = 5055
        sTab(0).Left = 5055
    Else
        lblPending.Top = 1080
        lblPending.Height = 375
        lblPending.Width = 2055
        lblPending.Caption = "Pending"
        sTab(2).Top = 1080
        sTab(2).Width = 2055
        sTab(2).FillColor = &H95075
        lblUpcoming.Left = 5055
        sTab(0).Left = 5055
        lblPending.Left = 7110
        sTab(2).Left = 7110
    End If
    TabSelected = 1
    Clear_All
    Load_All ("Archived")
End Sub

Private Sub lblSubject_Change(Index As Integer)
    On Error Resume Next
    If Vchanged = False Then
        If lblSubject(Index).Width > 3780 Then
            lblSubject(Index).ToolTipText = lblSubject(Index).Caption
            lblWork(Index).ToolTipText = lblSubject(Index).Caption
            Do
                Vchanged = True
                lblSubject(Index).Caption = Mid(lblSubject(Index).Caption, 1, Len(lblSubject(Index).Caption) - 1)
            Loop Until lblSubject(Index).Width <= 3720
            DoEvents
            Vchanged = True
            lblSubject(Index).Caption = lblSubject(Index).Caption + "."
            Vchanged = True
            lblSubject(Index).Caption = lblSubject(Index).Caption + "."
            Vchanged = True
            lblSubject(Index).Caption = lblSubject(Index).Caption + "."
        Else
            lblSubject(Index).ToolTipText = ""
        End If
    Else
        Vchanged = False
    End If
End Sub

Private Sub lblSubject_DblClick(Index As Integer)
    On Error Resume Next
    Call lblWork_DblClick(Index)
End Sub

Private Sub lblTime_Click(Index As Integer)
    On Error Resume Next
    Call lblWork_Click(Index)
End Sub
Public Sub lblDelete_Click()
    On Error Resume Next
    If WorksLoaded < 1 Then
        Exit Sub
    End If
    Dim tempstr As String
    Select Case TabSelected
        Case 0
            tempstr = "Upcoming"
        Case 1
            tempstr = "Archived"
        Case 2
            tempstr = "Pending"
    End Select
    DeleteSetting tempstr, Trim(Str(WorkSelected))
    If WorksLoaded > 1 Then
        For temp = WorkSelected To WorksLoaded - 2
            SaveSetting tempstr, Trim(Str(temp)), "Subject", GetSetting(tempstr, Trim(Str(temp + 1)), "Subject", "")
            lblSubject(temp).Caption = lblSubject(temp + 1).Caption
            If TabSelected <> 2 Then
                SaveSetting tempstr, Trim(Str(temp)), "Time", GetSetting(tempstr, Trim(Str(temp + 1)), "Time", "")
                lblTime(temp).Caption = lblTime(temp + 1).Caption
            End If
            SaveSetting tempstr, Trim(Str(temp)), "Content", GetSetting(tempstr, Trim(Str(temp + 1)), "Content", "")
        Next
    End If
    lblDelete.ForeColor = vbBlack
    SaveSetting tempstr, GetSetting(tempstr, "Total", "Number", "0"), "Temp", ""
    DeleteSetting tempstr, GetSetting(tempstr, "Total", "Number", "0")
    SaveSetting tempstr, "Total", "Number", Trim(Str(Val(GetSetting(tempstr, "Total", "Number", "0")) - 1))
    CalledExt = True
    WorkSelected = IIf(WorkSelected > 0, WorkSelected - 1, 0)
    Clear_All
    Load_All (tempstr)
    Call ContentFrame_MouseMove(0, 0, 0, 0)
End Sub
Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblDelete.ForeColor = vbWhite
End Sub
Private Sub lblDone_Click()
    On Error Resume Next
    If TabSelected = 2 Then
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Subject", lblSubject(WorkSelected).Caption
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Content", GetSetting("Pending", Trim(Str(WorkSelected)), "Content", "")
    Else
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Subject", lblSubject(WorkSelected).Caption
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Content", GetSetting("Upcoming", Trim(Str(WorkSelected)), "Content", "")
        SaveSetting "Archived", GetSetting("Archived", "Total", "Number", "0"), "Time", lblTime(WorkSelected).Caption
    End If
    SaveSetting "Archived", "Total", "Number", Trim(Str(Val(GetSetting("Archived", "Total", "Number", "0")) + 1))
    Call lblDelete_Click
    lblDone.ForeColor = vbBlack
End Sub
Private Sub lblDone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblDone.ForeColor = vbWhite
End Sub
Private Sub lblPending_Click()
    On Error Resume Next
    If CalledExt = False And TabSelected = 2 Then
        Exit Sub
    End If
    If TabSelected = 0 Then
        lblUpcoming.Top = 1080
        lblUpcoming.Height = 375
        lblUpcoming.Width = 2055
        lblUpcoming.Caption = "Upcoming"
        sTab(0).Top = 1080
        sTab(0).Width = 2055
        sTab(0).FillColor = &H95075
    Else
        lblArchived.Top = 1080
        lblArchived.Height = 375
        lblArchived.Width = 2055
        lblArchived.Caption = "Archived"
        sTab(1).Top = 1080
        sTab(1).Width = 2055
        sTab(1).FillColor = &H95075
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
    sTab(2).FillColor = &H125D82
    TabSelected = 2
    Clear_All
    Load_All ("Pending")
End Sub
Private Sub lblScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblScroll.BackColor = RGB(43, 30, 4)
    If Button = 1 Then
        Yinit = Y
        Mousedwn = True
    End If
End Sub
Private Sub lblScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblScroll.BackColor = RGB(174, 133, 49)
    If Button = 1 And Mousedwn = True Then
        lblScroll.Move lblScroll.Left, lblScroll.Top + Y - Yinit
        lblScroll.BackColor = RGB(43, 30, 4)
        For temp = WorksLoaded - 1 To 0 Step -1
            lblWork(temp).Top = lblWork(temp).Top - ((Y - Yinit) * ((720 * (WorksLoaded - 5)) / 2880))
            If TabSelected <> 2 Then
                lblTime(temp).Top = lblWork(temp).Top + 180
            End If
            lblSubject(temp).Top = lblWork(temp).Top + 120
        Next
        If lblDelete.Visible = True Then
            lblDelete.Top = lblWork(WorkSelected).Top + 180
        End If
        If lblDone.Visible = True Then
            lblDone.Top = lblWork(WorkSelected).Top + 180
        End If
    End If
End Sub

Private Sub lblScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblScroll.BackColor = RGB(154, 113, 29)
    If lblWork(WorksLoaded - 1).Top > 120 Then
        lblScroll.Top = 120
        lblWork(WorksLoaded - 1).Top = 120
        If TabSelected <> 2 Then
            lblTime(WorksLoaded - 1).Top = 300
        End If
        lblSubject(WorksLoaded - 1).Top = 240
        For temp = WorksLoaded - 2 To 0 Step -1
            lblWork(temp).Top = lblWork(temp + 1).Top + 720
            If TabSelected <> 2 Then
                lblTime(temp).Top = lblWork(temp).Top + 180
            End If
            lblSubject(temp).Top = lblWork(temp).Top + 120
        Next
    ElseIf lblWork(0).Top < 3000 Then
        lblScroll.Top = 3000
        lblWork(0).Top = 3000
        If TabSelected <> 2 Then
            lblTime(0).Top = 3180
        End If
        lblSubject(0).Top = 3120
        For temp = 1 To WorksLoaded - 1
            lblWork(temp).Top = lblWork(temp - 1).Top - 720
            If TabSelected <> 2 Then
                lblTime(temp).Top = lblWork(temp).Top + 180
            End If
            lblSubject(temp).Top = lblWork(temp).Top + 120
        Next
    End If
    If lblDelete.Visible = True Then
        lblDelete.Top = lblWork(WorkSelected).Top + 180
    End If
    If lblDone.Visible = True Then
        lblDone.Top = lblWork(WorkSelected).Top + 180
    End If
End Sub

Private Sub lblSubject_Click(Index As Integer)
    On Error Resume Next
    Call lblWork_Click(Index)
End Sub

Private Sub lblTime_DblClick(Index As Integer)
    On Error Resume Next
    Call lblWork_DblClick(Index)
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
    If Button = 1 And Mousedwn = True Then
        Move Left + X - Xinit, Top + Y - Yinit
    End If
End Sub

Private Sub lblUpcoming_Click()
    On Error Resume Next
    If CalledExt = False And TabSelected = 0 Then
        Exit Sub
    End If
    If TabSelected = 1 Then
        lblArchived.Top = 1080
        lblArchived.Height = 375
        lblArchived.Width = 2055
        lblArchived.Caption = "Archived"
        sTab(1).Top = 1080
        sTab(1).Width = 2055
        sTab(1).FillColor = &H95075
    Else
        lblPending.Top = 1080
        lblPending.Height = 375
        lblPending.Width = 2055
        lblPending.Caption = "Pending"
        sTab(2).Top = 1080
        sTab(2).Width = 2055
        sTab(2).FillColor = &H95075
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
    sTab(0).FillColor = &H125D82
    lblPending.Left = 7110
    sTab(2).Left = 7110
    TabSelected = 0
    Clear_All
    Load_All ("Upcoming")
End Sub

Private Sub lblWork_Click(Index As Integer)
    On Error Resume Next
    lblWork(WorkSelected).BackColor = &H1778A9
    lblWork(Index).BackColor = RGB(225, 165, 38)
    WorkSelected = Index
    If lblWork(WorkSelected).Top < 120 Then
        lblWork(WorkSelected).Top = 120
        lblTime(WorkSelected).Top = 300
        lblSubject(WorkSelected).Top = 240
        For temp = WorkSelected + 1 To WorksLoaded - 1
            lblWork(temp).Top = lblWork(temp - 1).Top + 720
            lblTime(temp).Top = lblWork(temp).Top + 180
            lblSubject(temp).Top = lblWork(temp).Top + 120
        Next
        If WorkSelected > 0 Then
            For temp = WorkSelected - 1 To 0 Step -1
                lblWork(temp).Top = lblWork(temp + 1).Top - 720
                lblTime(temp).Top = lblWork(temp).Top + 180
                lblSubject(temp).Top = lblWork(temp).Top + 120
            Next
        End If
        lblScroll.Top = 3000 - (lblWork(WorksLoaded - 5).Top - lblWork(WorkSelected).Top) * 2880 / (720 * (WorksLoaded - 5))
    End If
    If lblWork(WorkSelected).Top > 3000 Then
        lblWork(WorkSelected).Top = 3000
        lblTime(WorkSelected).Top = 3180
        lblSubject(WorkSelected).Top = 3120
        If WorkSelected < WorksLoaded - 1 Then
            For temp = WorkSelected + 1 To WorksLoaded - 1
                lblWork(temp).Top = lblWork(temp - 1).Top + 720
                lblTime(temp).Top = lblWork(temp).Top + 180
                lblSubject(temp).Top = lblWork(temp).Top + 120
            Next
        End If
        For temp = WorkSelected - 1 To 0 Step -1
            lblWork(temp).Top = lblWork(temp + 1).Top - 720
            lblTime(temp).Top = lblWork(temp).Top + 180
            lblSubject(temp).Top = lblWork(temp).Top + 120
        Next
        lblScroll.Top = 120 + (lblWork(WorkSelected).Top - lblWork(4).Top) * 2880 / (720 * (WorksLoaded - 5))
    End If
    lblDelete.Top = lblWork(WorkSelected).Top + 180
    lblDone.Top = lblWork(WorkSelected).Top + 180
    lblDelete.Visible = True
    If TabSelected <> 1 Then
        lblDone.Visible = True
    End If
End Sub
Private Sub lblWork_DblClick(Index As Integer)
    On Error Resume Next
    Load WINtask
    WINtask.Tag = "Edit"
    WINtask.Left = Left + (Width - WINtask.Width) / 2
    WINtask.Top = Top + (Height - WINtask.Height) / 2
    If TabSelected <> 2 Then
        If lblTime(Index).Caption <> "" Then
            WINtask.txtDeadline.Text = DateValue(lblTime(Index).Caption)
            WINtask.txtTime.Text = TimeValue(lblTime(Index).Caption)
        End If
    End If
    Select Case TabSelected
        Case 0
            WINtask.txtSubject.Text = GetSetting("Upcoming", Trim(Str(Index)), "Subject", "")
            WINtask.txtContent.Text = GetSetting("Upcoming", Trim(Str(Index)), "Content", "")
        Case 1
            WINtask.txtSubject.Text = GetSetting("Archived", Trim(Str(Index)), "Subject", "")
            WINtask.txtContent.Text = GetSetting("Archived", Trim(Str(Index)), "Content", "")
        Case 2
            WINtask.txtSubject.Text = GetSetting("Pending", Trim(Str(Index)), "Subject", "")
            WINtask.txtContent.Text = GetSetting("Pending", Trim(Str(Index)), "Content", "")
    End Select
    WINtask.Show vbModal
    If WINtask.Tag = "Editted" Then
        'Select Case TabSelected
            'Case 0
            '    SaveSetting "Upcoming", Trim(Str(WorkSelected)), "Subject", WINtask.txtSubject.Text
            '    SaveSetting "Upcoming", Trim(Str(WorkSelected)), "Content", WINtask.txtContent.Text
            '    SaveSetting "Upcoming", Trim(Str(WorkSelected)), "Time", WINtask.txtDeadline.Text
            '    lblSubject(WorkSelected).Caption = WINtask.txtSubject.Text
            '    lblTime(WorkSelected).Caption = WINtask.txtDeadline.Text
            'Case 1
            '    SaveSetting "Archived", Trim(Str(WorkSelected)), "Subject", WINtask.txtSubject.Text
            '    SaveSetting "Archived", Trim(Str(WorkSelected)), "Content", WINtask.txtContent.Text
            '    SaveSetting "Archived", Trim(Str(WorkSelected)), "Time", WINtask.txtDeadline.Text
            '    lblSubject(WorkSelected).Caption = WINtask.txtSubject.Text
            '   lblTime(WorkSelected).Caption = WINtask.txtDeadline.Text
            'Case 2
            '    SaveSetting "Pending", Trim(Str(WorkSelected)), "Subject", WINtask.txtSubject.Text
            '    SaveSetting "Pending", Trim(Str(WorkSelected)), "Content", WINtask.txtContent.Text
            '    lblSubject(WorkSelected).Caption = WINtask.txtSubject.Text
        'End Select
        Call lblDelete_Click
        Call AddTask(WINtask.txtSubject.Text, WINtask.txtContent.Text, WINtask.txtDeadline.Text)
    End If
    Unload WINtask
End Sub
Private Sub lblWork_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lblDelete.ForeColor = vbBlack
    lblDone.ForeColor = vbBlack
End Sub
Public Sub AddTask(s_Subject As String, s_Content As String, Optional s_Date As String = "")
    On Error Resume Next
    If s_Subject <> "" Then
        If s_Date = "" Then
            SaveSetting "Pending", GetSetting("Pending", "Total", "Number", "0"), "Subject", s_Subject
            SaveSetting "Pending", GetSetting("Pending", "Total", "Number", "0"), "Content", s_Content
            SaveSetting "Pending", "Total", "Number", Trim(Str(Val(GetSetting("Pending", "Total", "Number", "0")) + 1))
            WorkSelected = Val(GetSetting("Pending", "Total", "Number", "0")) - 1
            CalledExt = True
            Call lblPending_Click
        Else
            temp2 = 0
            If Val(GetSetting("Upcoming", "Total", "Number", "0")) > 0 Then
                For temp2 = 0 To Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1
                    If CDate(s_Date) > CDate(GetSetting("Upcoming", Trim(Str(temp2)), "Time")) Then
                        Exit For
                    End If
                Next
                For temp3 = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To temp2 Step -1
                    SaveSetting "Upcoming", Trim(Str(temp3 + 1)), "Subject", GetSetting("Upcoming", Trim(Str(temp3)), "Subject", "")
                    SaveSetting "Upcoming", Trim(Str(temp3 + 1)), "Time", GetSetting("Upcoming", Trim(Str(temp3)), "Time", "")
                    SaveSetting "Upcoming", Trim(Str(temp3 + 1)), "Content", GetSetting("Upcoming", Trim(Str(temp3)), "Content", "")
                Next
            End If
            SaveSetting "Upcoming", Trim(Str(temp2)), "Subject", s_Subject
            SaveSetting "Upcoming", Trim(Str(temp2)), "Content", s_Content
            SaveSetting "Upcoming", Trim(Str(temp2)), "Time", s_Date
            SaveSetting "Upcoming", "Total", "Number", Trim(Str(Val(GetSetting("Upcoming", "Total", "Number", "0")) + 1))
            WorkSelected = temp2
            CalledExt = True
            Call lblUpcoming_Click
        End If
    End If
End Sub
Public Sub Clear_All()
    On Error Resume Next
    For temp2 = lblWork.lBound + 1 To lblWork.UBound
        Unload lblWork(temp2)
    Next
    For temp2 = lblTime.lBound + 1 To lblTime.UBound
        Unload lblTime(temp2)
    Next
    For temp2 = lblSubject.lBound + 1 To lblSubject.UBound
        Unload lblSubject(temp2)
    Next
    lblDelete.Visible = False
    lblDone.Visible = False
    lblWork(0).Visible = False
    lblSubject(0).Visible = False
    lblTime(0).Visible = False
    WorksLoaded = 0
End Sub
Public Sub Load_All(sName As String)
    On Error Resume Next
    lblWork(0).BackColor = &H1778A9
    If Val(GetSetting(sName, "Total", "Number", "0")) = 0 Then
        lblScroll.Visible = False
        Exit Sub
    End If
    For temp3 = 1 To Val(GetSetting(sName, "Total", "Number", "0")) - 1
        Load lblWork(temp3)
        If sName <> "Pending" Then
            Load lblTime(temp3)
        End If
        Load lblSubject(temp3)
    Next
    lblSubject(temp3 - 1).Caption = GetSetting(sName, Trim(Str(Val(GetSetting(sName, "Total", "Number", "0")) - 1)), "Subject", "")
    lblWork(temp3 - 1).Visible = True
    lblSubject(temp3 - 1).Visible = True
    If sName <> "Pending" Then
        lblTime(temp3 - 1).Caption = GetSetting(sName, Trim(Str(Val(GetSetting(sName, "Total", "Number", "0")) - 1)), "Time", "")
        lblTime(temp3 - 1).Visible = True
        lblTime(temp3 - 1).Left = 240
        lblTime(temp3 - 1).ZOrder 0
    End If
    lblWork(temp3 - 1).Left = 120
    lblSubject(temp3 - 1).Left = 3120
    lblSubject(temp3 - 1).ZOrder 0
    For temp3 = temp3 - 2 To 0 Step -1
        lblSubject(temp3).Caption = GetSetting(sName, Trim(Str(temp3)), "Subject", "")
        lblWork(temp3).Visible = True
        lblSubject(temp3).Visible = True
        If sName <> "Pending" Then
            lblTime(temp3).Caption = GetSetting(sName, Trim(Str(temp3)), "Time", "")
            lblTime(temp3).Visible = True
            lblTime(temp3).ZOrder 0
        End If
        lblSubject(temp3).ZOrder 0
    Next
    WorksLoaded = Val(GetSetting(sName, "Total", "Number", "0"))
    If WorksLoaded > 5 Then
        lblScroll.Visible = True
    Else
        lblScroll.Visible = False
    End If
    lblScroll.Top = 120
    If CalledExt = False Then
        WorkSelected = WorksLoaded - 1
    End If
    If WorkSelected > WorksLoaded - 1 Then
        WorkSelected = WorksLoaded - 1
    ElseIf WorkSelected < 0 Then
        WorkSelected = 0
    End If
    lblWork(WorkSelected).BackColor = RGB(225, 165, 38)
    If WorksLoaded < 6 Then
        lblWork(WorkSelected).Top = 120 + (720 * ((WorksLoaded - 1) - WorkSelected))
    Else
        lblWork(WorkSelected).Top = 3000 - (720 * WorkSelected)
        If lblWork(WorkSelected).Top < 120 Then
            lblWork(WorkSelected).Top = 120
        End If
        If WorkSelected > 0 Then
            Call Form_KeyDown(vbKeyDown, 0)
            Call Form_KeyDown(vbKeyUp, 0)
        Else
            Call Form_KeyDown(vbKeyEnd, 0)
        End If
    End If
    For temp3 = 0 To WorksLoaded - 1
        lblWork(temp3).Top = lblWork(WorkSelected).Top + (720 * (WorkSelected - temp3))
        lblSubject(temp3).Top = lblWork(temp3).Top + 120
        If sName <> "Pending" Then
            lblTime(temp3).Top = lblWork(temp3).Top + 180
        End If
    Next
    lblDelete.Top = lblWork(WorkSelected).Top + 180
    lblDone.Top = lblDelete.Top
    lblDelete.Visible = True
    If sName <> "Archived" Then
        lblDone.Visible = True
    End If
    Call ContentFrame_MouseMove(0, 0, 0, 0)
    CalledExt = False
End Sub

Private Sub Notify_Timer()
    On Error Resume Next
    Dim nEvents As Integer
    Dim nEvents2 As Integer
    nEvents2 = 0
    
    'Notification about upcoming tasks in next hour
    
    If Val(Now - CDate(GetSetting("remindME", "Notify", "Latest Check", "01-01-01 00:00"))) > OneHr Then
        SaveSetting "remindME", "Notify", "Latest Check", Str(Now)
        If Val(GetSetting("Upcoming", "Total", "Number", "0")) > 0 Then
            Load WINnotify
            For nEvents = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To 0 Step -1
                If Val(CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) - Now) > OneHr Then
                    Exit For
                ElseIf Val(CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) - Now) < OneHr And Val(CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) - Now) > 0 Then
                    If GetSetting("Upcoming", Trim(Str(nEvents)), "Clicked", "0") = "0" Then
                        Load WINnotify
                        WINnotify.lblContent.Caption = WINnotify.lblContent.Caption + IIf(Val(CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) - Now) < 0, "(Missed) ", "") + GetSetting("Upcoming", Trim(Str(nEvents)), "Subject", "") + Chr(13)
                    End If
                    nEvents2 = nEvents2 + 1
                End If
            Next
            If nEvents2 > 0 Then
                WINnotify.lblSubject.Caption = Trim(Str(nEvents2)) + " upcoming " + "task" + IIf(nEvents2 <> 1, "s", "") + " in next one hour"
                WINnotify.Tag = ""
                WINnotify.Show
            End If
        End If
    End If
    If TabSelected <> 2 Then
        For nEvents = WorksLoaded - 1 To 0 Step -1
            If lblTime(nEvents).Caption <> "" Then
                If CDate(lblTime(nEvents).Caption) < Now Then
                    lblTime(nEvents).ForeColor = RGB(150, 0, 0)
                Else
                    lblTime(nEvents).ForeColor = vbBlack
                End If
            End If
        Next
    End If
    For nEvents = Val(GetSetting("Upcoming", "Total", "Number", "0")) - 1 To 0 Step -1
        If Now - CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) < OneSec And Now - CDate(GetSetting("Upcoming", Trim(Str(nEvents)), "Time", "")) > 0 And GetSetting("Upcoming", Trim(Str(nEvents)), "Clicked", "0") = "0" Then
            Load WINnotify
            WINnotify.lblContent.Caption = GetSetting("Upcoming", Trim(Str(nEvents)), "Content", "")
            WINnotify.lblSubject.Caption = GetSetting("Upcoming", Trim(Str(nEvents)), "Subject", "")
            WINnotify.lblContent.Top = WINnotify.lblSubject.Top + WINnotify.lblSubject.Height + 105
            If WINnotify.lblContent.Height + 105 + WINnotify.lblContent.Top > 2280 Then
                WINnotify.Height = WINnotify.lblContent.Height + WINnotify.lblContent.Top + 690
                Dim TempR As Integer
                For TempR = 0 To 5
                    WINnotify.lblPostpone(TempR).Top = WINnotify.Height - 465
                Next
            End If
            WINnotify.Tag = Trim(Str(nEvents))
            WINnotify.Show
        End If
    Next
End Sub
