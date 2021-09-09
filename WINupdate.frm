VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5BE2A8DD-6B60-4455-97ED-219C6250159F}#1.0#0"; "NetGrab.ocx"
Begin VB.Form WINupdate 
   Appearance      =   0  'Flat
   BackColor       =   &H0026A5E1&
   BorderStyle     =   0  'None
   Caption         =   "ANIco.in Updater Application"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ForeColor       =   &H0026A5E1&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "m"
   Begin NetGrabOCX.NetGrab NetGrabber 
      Left            =   5040
      Top             =   3120
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin VB.TextBox txtSave 
      Appearance      =   0  'Flat
      BackColor       =   &H0036B5F1&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   300
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   2400
   End
   Begin VB.TextBox txtVhist 
      Appearance      =   0  'Flat
      BackColor       =   &H0036B5F1&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   2100
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   4200
   End
   Begin VB.Timer TmrUpdate 
      Interval        =   1
      Left            =   4080
      Top             =   4080
   End
   Begin VB.Frame frmProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   4200
      Begin VB.Shape ShpJunk 
         BorderColor     =   &H80000006&
         Height          =   300
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   4200
      End
      Begin VB.Shape ShpProgress 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0046C5FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   0
         Top             =   0
         Width           =   900
      End
   End
   Begin MSWinsockLib.Winsock WinSock1 
      Left            =   4440
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label cmdStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0036B5F1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Start Download"
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
      Height          =   390
      Left            =   5400
      TabIndex        =   29
      Tag             =   "o"
      Top             =   4320
      Width           =   2445
   End
   Begin VB.Label optInstall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Do not Download Updates"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   2
      Left            =   5400
      TabIndex        =   28
      Top             =   3870
      Width           =   2400
   End
   Begin VB.Label lblTick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   300
      Index           =   3
      Left            =   5160
      TabIndex        =   26
      Top             =   3315
      Width           =   240
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Download & Installation"
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
      Index           =   6
      Left            =   5280
      TabIndex        =   25
      Top             =   3000
      UseMnemonic     =   0   'False
      Width           =   2085
   End
   Begin VB.Label optInstall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Download And Install &Later"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   1
      Left            =   5400
      TabIndex        =   24
      Top             =   3615
      Width           =   2400
   End
   Begin VB.Label optInstall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Download And &Install Now"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   23
      Top             =   3360
      Width           =   2400
   End
   Begin VB.Label lblTick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   300
      Index           =   2
      Left            =   5160
      TabIndex        =   22
      Top             =   1635
      Width           =   240
   End
   Begin VB.Label lblTick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   300
      Index           =   1
      Left            =   5160
      TabIndex        =   21
      Top             =   675
      Width           =   240
   End
   Begin VB.Label lblTick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   4410
      Width           =   240
   End
   Begin VB.Line lnJunk 
      BorderColor     =   &H80000006&
      X1              =   4920
      X2              =   4920
      Y1              =   5280
      Y2              =   240
   End
   Begin VB.Label optSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "B&rowse Manually"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   2
      Left            =   5400
      TabIndex        =   18
      Top             =   2190
      Width           =   2400
   End
   Begin VB.Label optSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Temporary Directory"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   17
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label optSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Downloads Directory"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   1
      Left            =   5400
      TabIndex        =   16
      Top             =   1935
      Width           =   2400
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Save to..."
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
      Index           =   5
      Left            =   5280
      TabIndex        =   15
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   780
   End
   Begin VB.Label optMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&ANIco.in Downloader"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   14
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label optMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Custom Download Manager"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   1
      Left            =   5400
      TabIndex        =   13
      Top             =   975
      Width           =   2400
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading Method"
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
      Index           =   4
      Left            =   5280
      TabIndex        =   12
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   1905
   End
   Begin VB.Label optAUN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Major Versions Only (Less Frequent)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   11
      ToolTipText     =   "eg. 1.4 -> 2.0"
      Top             =   4200
      Width           =   4080
   End
   Begin VB.Label optAUN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mi&nor Versions (Recommended)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   10
      ToolTipText     =   "eg. 1.4 -> 1.5"
      Top             =   4455
      Width           =   4080
   End
   Begin VB.Label optAUN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Build Versions (More Frequent)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "eg. 1.4.108 -> 1.4.109"
      Top             =   4710
      Width           =   4080
   End
   Begin VB.Label optAUN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N&o Update Notification (Manual Only)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   4965
      Width           =   4080
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version History"
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
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic Updates Notification"
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
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      UseMnemonic     =   0   'False
      Width           =   2730
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Checking for updates..."
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ANIco.in"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   255
      Index           =   1
      Left            =   5550
      TabIndex        =   2
      Top             =   4815
      Width           =   825
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATER BETA 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   150
      Index           =   3
      Left            =   5520
      TabIndex        =   1
      Top             =   5010
      Width           =   885
   End
   Begin VB.Shape ShpJunk 
      BorderColor     =   &H80000006&
      Height          =   390
      Index           =   1
      Left            =   5400
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Label cmdHide 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0036B5F1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hid&e"
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
      Height          =   390
      Left            =   6720
      TabIndex        =   0
      Tag             =   "o"
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   5520
      Index           =   10
      Left            =   0
      TabIndex        =   27
      Tag             =   "o"
      Top             =   0
      Width           =   8160
   End
End
Attribute VB_Name = "WINupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AUNsel As Integer
Dim SaveSel As Integer
Dim InstallSel As Integer
Dim MethodSel As Integer
Dim ProgressDirection As Integer
Dim UpdateSize As Double
Dim UpdaterStage As Byte
Dim InitBytes As Long
Dim DLbytes As Long
Dim InitTmr As Single

Private Sub cmdHide_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    If MethodSel = 0 Then
        UpdaterStage = 1
        TmrUpdate.Enabled = True
        InitTmr = Timer
        NetGrabber.DownloadStart Trim$("http://dl.dropbox.com/u/71359423/Executables/" & App.ProductName & ".exe"), vbAsyncReadForceUpdate
        UpdateSize = Val(Mid$(txtVhist.Text, GetNthOcc(1, txtVhist.Text, "(") + 1, GetNthOcc(2, txtVhist.Text, " ") - GetNthOcc(1, txtVhist.Text, "(") - 1))
    Else
        ShellExecute 0, "open", Chr(34) + "http://dl.dropbox.com/u/71359423/Executables/" & App.ProductName & ".exe" + Chr(34), 0, 0, 1
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If WinSock1.LocalIP = "127.0.0.1" Or WinSock1.LocalIP = vbNullString Then
        Dim objControl As Control
        For Each objControl In Me.Controls
            If objControl.Name <> "cmdHide" And objControl.Name <> "txtVhist" Then
                objControl.Enabled = False
            End If
        Next
        txtVhist.Text = "NO INTERNET ACCESS" & vbCrLf & vbCrLf & "Try Again later"
    End If
    AUNsel = 1
    ProgressDirection = 1
    NetGrabber.DownloadStart Trim$("http://dl.dropbox.com/u/71359423/Update/" & App.ProductName & ".txt"), vbAsyncReadForceUpdate
End Sub

Private Sub NetGrabber_DownloadComplete(ByVal nBytes As Long)
    If UpdaterStage = 0 Then
        txtVhist.Text = StrConv(NetGrabber.Bytes, vbUnicode)
        ShpProgress.Width = 15
        TmrUpdate.Enabled = False
        ShpProgress.Left = 0
        TmrUpdate.Interval = 1000
        If Tag = vbNullString Then                          'If updates were checked without user consent
        
                'Application's major update available           Application's minor update is available                                                                                                                                     Application's build (revision) update is available
            If (Val(Mid$(txtVhist.Text, 1, 1)) > App.Major) Or ((Val(Mid$(txtVhist.Text, GetNthOcc(1, txtVhist.Text, ".") + 1, GetNthOcc(2, txtVhist.Text, ".") - GetNthOcc(1, txtVhist.Text, ".") - 1)) > App.Minor) And AUNsel > 0) Or ((Val(Mid$(txtVhist.Text, GetNthOcc(2, txtVhist.Text, ".") + 1, GetNthOcc(1, txtVhist.Text, " ") - GetNthOcc(2, txtVhist.Text, ".") - 1)) > App.Revision) And AUNsel > 1) Then
                lblStatus.Caption = "Updates are ready for download"
                cmdStart.Enabled = True
                Show
            Else
                Unload Me
            End If
        Else
            If (Val(Mid$(txtVhist.Text, 1, 1)) > App.Major) Or (Val(Mid$(txtVhist.Text, GetNthOcc(1, txtVhist.Text, ".") + 1, GetNthOcc(2, txtVhist.Text, ".") - GetNthOcc(1, txtVhist.Text, ".") - 1)) > App.Minor) Or (Val(Mid$(txtVhist.Text, GetNthOcc(2, txtVhist.Text, ".") + 1, GetNthOcc(1, txtVhist.Text, " ") - GetNthOcc(2, txtVhist.Text, ".") - 1)) > App.Revision) Then
                cmdStart.Enabled = True
                lblStatus.Caption = "Updates are ready for download"
            Else
                lblStatus.Caption = "Running latest version"
            End If
        End If
    ElseIf UpdaterStage = 1 Then
        Dim DownloadDir As String
        Select Case SaveSel
            Case 0
                DownloadDir = Environ("TEMP") & "\"
            Case 1
                DownloadDir = Environ("USERPROFILE") & "\Downloads\"
            Case 2
                DownloadDir = txtSave.Text & "\"
        End Select
        TmrUpdate.Enabled = False
        lblStatus.Caption = "Downloaded " & SizeString(UpdateSize) & " @" & SizeString(nBytes / IIf(Timer - InitTmr > 0, Timer - InitTmr, Timer + 86400 - InitTmr)) & "ps (Average)"
        ShpProgress.Width = 4200
        NetGrabber.SaveAs DownloadDir & App.ProductName & ".exe"
    End If
End Sub

Private Sub NetGrabber_DownloadProgress(ByVal nBytes As Long)
    If UpdaterStage = 1 Then
        DLbytes = nBytes
    End If
End Sub

Private Sub optAUN_Click(Index As Integer)
    optAUN(AUNsel).FontBold = False
    optAUN(Index).FontBold = True
    lblTick(0).Top = optAUN(Index).Top - 45
    AUNsel = Index
End Sub

Private Sub optInstall_Click(Index As Integer)
    optInstall(InstallSel).FontBold = False
    optInstall(Index).FontBold = True
    lblTick(3).Top = optInstall(Index).Top - 45
    InstallSel = Index
End Sub

Private Sub optMethod_Click(Index As Integer)
    optMethod(MethodSel).FontBold = False
    optMethod(Index).FontBold = True
    lblTick(1).Top = optMethod(Index).Top - 45
    If Index = 0 Then
        EnableDownloader
    Else
        DisableDownloader
    End If
    MethodSel = Index
End Sub

Private Sub EnableDownloader()
    lblJunk(5).Enabled = True
    lblJunk(6).Enabled = True
    optSave(0).Enabled = True
    optSave(1).Enabled = True
    optSave(2).Enabled = True
    optInstall(0).Enabled = True
    optInstall(1).Enabled = True
    optInstall(2).Enabled = True
    lblTick(2).Enabled = True
    lblTick(3).Enabled = True
    txtSave.Enabled = True
End Sub

Private Sub DisableDownloader()
    lblJunk(5).Enabled = False
    lblJunk(6).Enabled = False
    optSave(0).Enabled = False
    optSave(1).Enabled = False
    optSave(2).Enabled = False
    optInstall(0).Enabled = False
    optInstall(1).Enabled = False
    optInstall(2).Enabled = False
    lblTick(2).Enabled = False
    lblTick(3).Enabled = False
    txtSave.Enabled = False
End Sub

Private Sub optSave_Click(Index As Integer)
    optSave(SaveSel).FontBold = False
    optSave(Index).FontBold = True
    lblTick(2).Top = optSave(Index).Top - 45
    SaveSel = Index
End Sub

Private Sub StartDownload()
    lblJunk(4).Enabled = True
    optMethod(0).Enabled = True
    optMethod(1).Enabled = True
    lblTick(1).Enabled = True
    If optMethod(0).FontBold Then
        EnableDownloader
    End If
End Sub

Private Sub TmrUpdate_Timer()
    Select Case UpdaterStage
        Case 0
            ShpProgress.Left = ShpProgress.Left + (45 * ProgressDirection)
            If (ShpProgress.Left + ShpProgress.Width >= ShpJunk(0).Width) Or ShpProgress.Left <= 0 Then
                ProgressDirection = -1 * ProgressDirection
            End If
        Case 1
            lblStatus.Caption = "Downloading " & SizeString(UpdateSize) & " @" & SizeString(DLbytes - InitBytes) & "ps (" & Trim$(Int(DLbytes * 100 / UpdateSize)) & "% Done)"
            ShpProgress.Width = Int(DLbytes * 280 / UpdateSize) * 15
            InitBytes = DLbytes
    End Select
End Sub

Private Function GetNthOcc(n As Double, SearchStr As String, FindStr As String) As Double
    Dim i As Double
    Dim FindCt As Double
    i = 1
    Do Until i = 0 Or FindCt = n
        i = InStr(i, SearchStr, FindStr)
        If i > 0 Then
            FindCt = FindCt + 1
            i = i + 1
        End If
    Loop
    GetNthOcc = i - 1
End Function

Private Function SizeString(SizeValue As Double) As String
    Select Case SizeValue
        Case Is < 1000
            SizeString = Trim$(Round(SizeValue, 2)) & " B"
        Case Is < 1000000
            SizeString = Trim$(Round((SizeValue / 1024), 2)) & " KB"
        Case Is < 1000000000
            SizeString = Trim$(Round(SizeValue / (1048576), 2)) & " MB"
        Case Else
            SizeString = Trim$(Round(SizeValue / (1073741824), 2)) & " GB"
    End Select
End Function
