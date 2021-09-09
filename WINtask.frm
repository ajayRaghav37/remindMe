VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form WINtask 
   BackColor       =   &H001778A9&
   BorderStyle     =   0  'None
   Caption         =   "Add/Edit Task"
   ClientHeight    =   3255
   ClientLeft      =   2715
   ClientTop       =   3360
   ClientWidth     =   6615
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDeadline 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin MSComCtl2.MonthView MyCalen 
      Height          =   2910
      Left            =   2400
      TabIndex        =   3
      Top             =   165
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
      StartOfWeek     =   30867458
      TitleBackColor  =   1538217
      TitleForeColor  =   0
      TrailingForeColor=   1538217
      CurrentDate     =   40885
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3120
      TabIndex        =   4
      Top             =   2655
      Width           =   975
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1965
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   615
      Width           =   5295
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5295
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
      Height          =   405
      Left            =   5400
      TabIndex        =   10
      Top             =   2670
      Width           =   1125
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001778A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&OK"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   2670
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   6615
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2685
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
      Left            =   120
      TabIndex        =   7
      Top             =   2685
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
      Left            =   120
      TabIndex        =   6
      Top             =   645
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
      Left            =   120
      TabIndex        =   5
      Top             =   165
      Width           =   870
   End
End
Attribute VB_Name = "WINtask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    On Error Resume Next
    txtSubject.Text = ""
    Hide
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdOK.BackColor = RGB(169, 120, 23)
    cmdCancel.BackColor = RGB(225, 165, 38)
End Sub
Private Sub cmdOK_Click()
    On Error Resume Next
    If txtTime.Text <> "" And txtDeadline.Text = "" Then
        txtDeadline.Text = DateValue(Now)
    End If
    If txtDeadline.Text <> "" And IsDate(txtDeadline.Text + " " + txtTime.Text) = False Then
        txtDeadline.ForeColor = vbRed
        txtTime.ForeColor = vbRed
        Exit Sub
    End If
    If txtDeadline.Text <> "" Then
        If (Now > CDate(txtDeadline.Text + " " + txtTime.Text)) Then
            If DateValue(txtDeadline.Text) = DateValue(Now) And txtTime.Text = "" Then
            Else
                txtDeadline.ForeColor = vbBlue
                txtTime.ForeColor = vbBlue
                Exit Sub
            End If
        End If
    End If
    If txtDeadline.Text <> "" Then
        txtDeadline.Text = CDate(txtDeadline.Text + "  " + txtTime.Text)
        txtTime.Text = ""
    End If
    Tag = "Editted"
    Hide
End Sub
Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdCancel.BackColor = RGB(169, 120, 23)
    cmdOK.BackColor = RGB(225, 165, 38)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            Call cmdCancel_Click
        Case vbKeyReturn
            If ActiveControl <> txtContent And ActiveControl <> MyCalen Then
                Call cmdOK_Click
            End If
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtSubject.BackColor = RGB(225, 165, 38)
    txtTime.BackColor = RGB(225, 165, 38)
    txtDeadline.BackColor = RGB(225, 165, 38)
    txtContent.BackColor = RGB(225, 165, 38)
    MyCalen.MinDate = Now
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    cmdOK.BackColor = RGB(169, 120, 23)
    cmdCancel.BackColor = RGB(169, 120, 23)
End Sub

Private Sub MyCalen_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    txtDeadline.Text = DateValue(CDate(DateClicked))
    MyCalen.Visible = False
    txtTime.SetFocus
End Sub

Private Sub MyCalen_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        Call MyCalen_DateClick(MyCalen.Value)
    End If
End Sub

Private Sub txtContent_GotFocus()
    On Error Resume Next
    MyCalen.Visible = False
End Sub

Private Sub txtDeadline_Change()
    On Error Resume Next
    txtDeadline.ForeColor = vbBlack
    txtTime.ForeColor = vbBlack
    MyCalen.Visible = True
End Sub

Private Sub txtDeadline_GotFocus()
    On Error Resume Next
    MyCalen.Visible = True
End Sub

Private Sub txtSubject_GotFocus()
    On Error Resume Next
    MyCalen.Visible = False
End Sub

Private Sub txtTime_Change()
    On Error Resume Next
    txtDeadline.ForeColor = vbBlack
    txtTime.ForeColor = vbBlack
End Sub

Private Sub txtTime_GotFocus()
    On Error Resume Next
    MyCalen.Visible = False
End Sub
