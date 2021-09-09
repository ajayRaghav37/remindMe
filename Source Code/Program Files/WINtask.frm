VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form WINtask 
   Appearance      =   0  'Flat
   BackColor       =   &H001778A9&
   BorderStyle     =   0  'None
   Caption         =   "Add/Edit Task"
   ClientHeight    =   3525
   ClientLeft      =   2715
   ClientTop       =   3360
   ClientWidth     =   6885
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView ExpiryCalen 
      Height          =   2910
      Left            =   3195
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   5133
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   2532833
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   2532833
      ShowToday       =   0   'False
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   41549826
      TitleBackColor  =   1538217
      TitleForeColor  =   0
      TrailingForeColor=   1538217
      CurrentDate     =   73050
      MaxDate         =   73050
   End
   Begin VB.TextBox txtExpiry 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      Enabled         =   0   'False
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
      Height          =   420
      Left            =   5505
      TabIndex        =   5
      Top             =   2985
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDeadline 
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
      Height          =   420
      Left            =   1440
      TabIndex        =   2
      Top             =   2985
      Width           =   1215
   End
   Begin MSComCtl2.MonthView MyCalen 
      Height          =   2910
      Left            =   1440
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   5133
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   2532833
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   2532833
      ShowToday       =   0   'False
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   41549826
      TitleBackColor  =   1538217
      TitleForeColor  =   0
      TrailingForeColor=   1538217
      CurrentDate     =   40885
      MaxDate         =   73050
   End
   Begin VB.TextBox txtTime 
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
      Height          =   420
      Left            =   3360
      TabIndex        =   4
      Top             =   2985
      Width           =   975
   End
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
      Height          =   1920
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   5280
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5280
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   4410
      TabIndex        =   19
      Top             =   2985
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occurrence"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   150
      TabIndex        =   18
      Top             =   2580
      UseMnemonic     =   0   'False
      Width           =   1185
   End
   Begin VB.Label FreqO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Once"
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
      Height          =   345
      Left            =   1440
      TabIndex        =   17
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label FreqD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&D'y"
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
      Left            =   2520
      TabIndex        =   16
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label FreqW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&W'y"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label FreqM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&M'y"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label FreqY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Y'y"
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
      Left            =   5760
      TabIndex        =   13
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Cancel"
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
      Height          =   420
      Left            =   5595
      TabIndex        =   12
      Top             =   2985
      Width           =   1125
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O&K"
      Enabled         =   0   'False
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
      Height          =   420
      Left            =   4410
      TabIndex        =   11
      Tag             =   "o"
      Top             =   2985
      Width           =   1125
   End
   Begin VB.Shape sBoundary 
      Height          =   3525
      Left            =   0
      Top             =   0
      Width           =   6885
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2760
      TabIndex        =   10
      Top             =   2985
      UseMnemonic     =   0   'False
      Width           =   510
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deadline"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   2985
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Content"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   825
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject*"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "WINtask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    On Error Resume Next
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    WINtaskLoaded = True
    MyCalen.MinDate = Now
    ExpiryCalen.MinDate = Now
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdOK.BackColor = cButton_Normal
    cmdCancel.BackColor = cButton_Normal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            cmdCancel_Click
        Case vbKeyReturn
            If ActiveControl <> txtContent And ActiveControl <> MyCalen Then
                cmdOK_Click
            End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyK
            If Shift = vbAltMask Then
                cmdOK_Click
            End If
        Case vbKeyC
            If Shift = vbAltMask Then
                cmdCancel_Click
            End If
        Case vbKeyO
            If Shift = vbAltMask Then
                FreqO_Click
            End If
        Case vbKeyD
            If Shift = vbAltMask Then
                FreqD_Click
            End If
        Case vbKeyW
            If Shift = vbAltMask Then
                FreqW_Click
            End If
        Case vbKeyM
            If Shift = vbAltMask Then
                FreqM_Click
            End If
        Case vbKeyY
            If Shift = vbAltMask Then
                FreqY_Click
            End If
    End Select
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    txtSubject.Text = ""
    Tag = ""
    Hide
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdOK.BackColor = cButton_Normal
    cmdCancel.BackColor = cButton_Hover
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    If txtTime.Text <> "" And txtDeadline.Text = "" Then
        txtDeadline.Text = DateValue(Now)
    End If
    If txtExpiry.Text = "" And cmdOK.Tag <> "o" Then
        txtExpiry.Text = Str(ExpiryCalen.Value)
    End If
    If txtDeadline.Text <> "" And Not IsDate(txtDeadline.Text + " " + txtTime.Text) Then
        txtDeadline.ForeColor = vbDarkRed
        txtTime.ForeColor = vbDarkRed
        MyCalen.Visible = False
        Exit Sub
    End If
    If Not IsDate(txtExpiry.Text) And cmdOK.Tag <> "o" Then
        txtExpiry.ForeColor = vbDarkRed
        ExpiryCalen.Visible = False
        Exit Sub
    End If
    If txtDeadline.Text <> "" Then
        If (Now > CDate(txtDeadline.Text + " " + txtTime.Text)) Then
            If DateValue(txtDeadline.Text) = DateValue(Now) And txtTime.Text = "" Then
            Else
                txtDeadline.ForeColor = vbBlue
                txtTime.ForeColor = vbBlue
                MyCalen.Visible = False
                Exit Sub
            End If
        End If
    End If
    If txtExpiry.Text <> "" And cmdOK.Tag <> "o" Then
        If (CDate(txtDeadline.Text) > CDate(txtExpiry.Text)) Then
            txtExpiry.ForeColor = vbBlue
            ExpiryCalen.Visible = False
            Exit Sub
        End If
    End If
    If txtDeadline.Text <> "" Then
        txtDeadline.Text = CDate(txtDeadline.Text + "  " + txtTime.Text)
        txtTime.Text = ""
    End If
    If txtExpiry.Text <> "" And cmdOK.Tag <> "o" Then
        txtExpiry.Text = CDate(txtExpiry.Text)
    End If
    Tag = "Editted"
    Hide
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdOK.BackColor = cButton_Hover
    cmdCancel.BackColor = cButton_Normal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WINtaskLoaded = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WINtaskLoaded = False
End Sub

Private Sub FreqD_Click()
    On Error Resume Next
    cmdOK.Tag = "d"
    FreqO.BackColor = cButton_Normal
    FreqO.ForeColor = vbGray
    FreqW.BackColor = cButton_Normal
    FreqW.ForeColor = vbGray
    FreqY.BackColor = cButton_Normal
    FreqY.ForeColor = vbGray
    FreqM.BackColor = cButton_Normal
    FreqM.ForeColor = vbGray
    FreqD.BackColor = cButton_Hover
    FreqD.ForeColor = vbBlack
    lblJunk(2).Caption = "From Date*"
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
    If txtDeadline.Text = "" Then
        cmdOK.Enabled = False
        txtDeadline.SetFocus
    End If
    ShowExpiry
End Sub

Private Sub FreqM_Click()
    On Error Resume Next
    cmdOK.Tag = "m"
    FreqO.BackColor = cButton_Normal
    FreqO.ForeColor = vbGray
    FreqW.BackColor = cButton_Normal
    FreqW.ForeColor = vbGray
    FreqY.BackColor = cButton_Normal
    FreqY.ForeColor = vbGray
    FreqD.BackColor = cButton_Normal
    FreqD.ForeColor = vbGray
    FreqM.BackColor = cButton_Hover
    FreqM.ForeColor = vbBlack
    lblJunk(2).Caption = "From Date*"
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
    If txtDeadline.Text = "" Then
        cmdOK.Enabled = False
        txtDeadline.SetFocus
    End If
    ShowExpiry
End Sub

Private Sub FreqO_Click()
    On Error Resume Next
    cmdOK.Tag = "o"
    FreqD.BackColor = cButton_Normal
    FreqD.ForeColor = vbGray
    FreqW.BackColor = cButton_Normal
    FreqW.ForeColor = vbGray
    FreqY.BackColor = cButton_Normal
    FreqY.ForeColor = vbGray
    FreqM.BackColor = cButton_Normal
    FreqM.ForeColor = vbGray
    FreqO.BackColor = cButton_Hover
    FreqO.ForeColor = vbBlack
    lblJunk(2).Caption = "Deadline"
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
    If txtSubject.Text <> "" Then
        cmdOK.Enabled = True
    End If
    HideExpiry
End Sub

Private Sub FreqW_Click()
    On Error Resume Next
    cmdOK.Tag = "ww"
    FreqO.BackColor = cButton_Normal
    FreqO.ForeColor = vbGray
    FreqD.BackColor = cButton_Normal
    FreqD.ForeColor = vbGray
    FreqY.BackColor = cButton_Normal
    FreqY.ForeColor = vbGray
    FreqM.BackColor = cButton_Normal
    FreqM.ForeColor = vbGray
    FreqW.BackColor = cButton_Hover
    FreqW.ForeColor = vbBlack
    lblJunk(2).Caption = "From Date*"
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
    If txtDeadline.Text = "" Then
        cmdOK.Enabled = False
        txtDeadline.SetFocus
    End If
    ShowExpiry
End Sub

Private Sub FreqY_Click()
    On Error Resume Next
    cmdOK.Tag = "yyyy"
    FreqO.BackColor = cButton_Normal
    FreqO.ForeColor = vbGray
    FreqW.BackColor = cButton_Normal
    FreqW.ForeColor = vbGray
    FreqD.BackColor = cButton_Normal
    FreqD.ForeColor = vbGray
    FreqM.BackColor = cButton_Normal
    FreqM.ForeColor = vbGray
    FreqY.BackColor = cButton_Hover
    FreqY.ForeColor = vbBlack
    lblJunk(2).Caption = "From Date*"
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
    If txtDeadline.Text = "" Then
        cmdOK.Enabled = False
        txtDeadline.SetFocus
    End If
    ShowExpiry
End Sub

Private Sub lblJunk_Click(Index As Integer)
    On Error Resume Next
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
End Sub

Private Sub MyCalen_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    txtDeadline.Text = DateClicked
    MyCalen.Visible = False
    ExpiryCalen.MinDate = DateClicked
    txtTime.SetFocus
End Sub

Private Sub MyCalen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        MyCalen_DateClick (MyCalen.Value)
    End If
End Sub

Private Sub ExpiryCalen_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    txtExpiry.Text = DateClicked
    MyCalen.MaxDate = DateClicked
    ExpiryCalen.Visible = False
End Sub

Private Sub ExpiryCalen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        ExpiryCalen_DateClick (ExpiryCalen.Value)
    End If
End Sub

Private Sub txtContent_GotFocus()
    On Error Resume Next
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
End Sub

Private Sub txtDeadline_Change()
    On Error Resume Next
    txtDeadline.ForeColor = vbBlack
    txtTime.ForeColor = vbBlack
    If txtSubject.Text <> "" Then
        cmdOK.Enabled = True
    End If
End Sub

Private Sub txtExpiry_Change()
    On Error Resume Next
    txtExpiry.ForeColor = vbBlack
End Sub

Private Sub txtDeadline_Click()
    On Error Resume Next
    ShowMyCalen
End Sub

Private Sub txtExpiry_Click()
    On Error Resume Next
    ShowExpiryCalen
End Sub

Private Sub txtDeadline_GotFocus()
    On Error Resume Next
    ShowMyCalen
    ExpiryCalen.Visible = False
End Sub

Private Sub txtExpiry_GotFocus()
    On Error Resume Next
    ShowExpiryCalen
    MyCalen.Visible = False
End Sub

Private Sub txtSubject_Change()
    On Error Resume Next
    If txtSubject.Text = "" Then
        cmdOK.Enabled = False
    Else
        If cmdOK.Tag = "o" Or cmdOK.Tag = "" Then
            cmdOK.Enabled = True
        Else
            If txtDeadline.Text <> "" Then
                cmdOK.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub txtSubject_GotFocus()
    On Error Resume Next
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
End Sub

Private Sub txtSubject_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtSubject.SelStart = 0
        txtSubject.SelLength = Len(txtSubject.Text)
    End If
End Sub

Private Sub txtDeadline_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtDeadline.SelStart = 0
        txtDeadline.SelLength = Len(txtDeadline.Text)
        ShowMyCalen
    End If
End Sub

Private Sub txtExpiry_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtExpiry.SelStart = 0
        txtExpiry.SelLength = Len(txtExpiry.Text)
        ShowExpiryCalen
    End If
End Sub

Private Sub txtTime_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtTime.SelStart = 0
        txtTime.SelLength = Len(txtTime.Text)
    End If
End Sub

Private Sub txtContent_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtContent.SelStart = 0
        txtContent.SelLength = Len(txtContent.Text)
    End If
End Sub

Private Sub txtTime_Change()
    On Error Resume Next
    txtDeadline.ForeColor = vbBlack
    txtTime.ForeColor = vbBlack
End Sub

Private Sub txtTime_GotFocus()
    On Error Resume Next
    MyCalen.Visible = False
    ExpiryCalen.Visible = False
End Sub

Public Sub ShowExpiry()
    On Error Resume Next
    If txtExpiry.Visible Then
        Exit Sub
    End If
    txtExpiry.Enabled = True
    txtExpiry.Visible = True
    ExpiryCalen.Enabled = True
    lblJunk(5).Visible = True
    cmdOK.Top = cmdOK.Top + 60 + txtExpiry.Height
    cmdCancel.Top = cmdOK.Top
    Height = Height + 60 + txtExpiry.Height
    sBoundary.Height = Height
End Sub

Private Sub HideExpiry()
    On Error Resume Next
    If Not txtExpiry.Visible Then
        Exit Sub
    End If
    txtExpiry.Enabled = False
    txtExpiry.Visible = False
    ExpiryCalen.Visible = False
    ExpiryCalen.Enabled = False
    lblJunk(5).Visible = False
    cmdOK.Top = cmdOK.Top - 60 - txtExpiry.Height
    cmdCancel.Top = cmdOK.Top
    Height = Height - 60 - txtExpiry.Height
    sBoundary.Height = Height
End Sub

Private Sub ShowMyCalen()
    On Error Resume Next
    MyCalen.Visible = True
    If txtDeadline.Text <> "" And IsDate(txtDeadline.Text) Then
        MyCalen.Value = txtDeadline.Text
    Else
        MyCalen.Value = Now
    End If
    If txtExpiry.Text <> "" And IsDate(txtExpiry.Text) Then
        MyCalen.MaxDate = txtExpiry.Text
    End If
End Sub

Private Sub ShowExpiryCalen()
    On Error Resume Next
    ExpiryCalen.Visible = True
    If txtExpiry.Text <> "" And IsDate(txtExpiry.Text) Then
        ExpiryCalen.Value = txtExpiry.Text
    Else
        If IsDate(txtDeadline.Text) Then
            ExpiryCalen.Value = DateAdd(cmdOK.Tag, 1, CDate(txtDeadline.Text))
        Else
            ExpiryCalen.Value = DateAdd(cmdOK.Tag, 1, Now)
        End If
    End If
    If txtDeadline.Text <> "" And IsDate(txtDeadline.Text) Then
        ExpiryCalen.MinDate = txtDeadline.Text
    End If
End Sub
