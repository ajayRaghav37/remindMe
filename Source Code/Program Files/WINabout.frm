VERSION 5.00
Begin VB.Form WINabout 
   Appearance      =   0  'Flat
   BackColor       =   &H0026A5E1&
   BorderStyle     =   0  'None
   Caption         =   "About remindME"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label cmdHide 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Hide"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3660
      TabIndex        =   7
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label cmdSource 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Download Source"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1740
      TabIndex        =   6
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Label cmdLicense 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View &License"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   4
      Left            =   330
      TabIndex        =   4
      Top             =   1560
      Width           =   4020
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2012 ANIco.in"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   4680
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Version: 1.0.32.4058"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4680
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackColor       =   &H0026A5E1&
      BackStyle       =   0  'Transparent
      Caption         =   "A reminder software by ANIco.in"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   4680
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      BackColor       =   &H0026A5E1&
      BackStyle       =   0  'Transparent
      Caption         =   "remindME"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4680
   End
   Begin VB.Shape shpBoundary 
      BorderColor     =   &H80000006&
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "WINabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHide_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdLicense_Click()
    On Error Resume Next
    ShellExecute 0, "open", Chr$(34) + "http://www.gnu.org/licenses/gpl.html" + Chr$(34), 0, 0, 1
End Sub

Private Sub cmdSource_Click()
    On Error Resume Next
    ShellExecute 0, "open", Chr$(34) + "http://dl.dropbox.com/u/71359423/Compressed/remindME.zip" + Chr$(34), 0, 0, 1
End Sub

Private Sub Form_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyD And Shift = vbAltMask Then
        cmdSource_Click
    End If
    If KeyCode = vbKeyL And Shift = vbAltMask Then
        cmdLicense_Click
    End If
    If KeyCode = vbKeyH And Shift = vbAltMask Then
        cmdHide_Click
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Top = WINremindME.Top + WINremindME.lblAbout.Top - Height
    Left = WINremindME.Left + WINremindME.lblAbout.Left - (Width - WINremindME.lblAbout.Width) / 2
    lblJunk(2).Caption = "Current Version: " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    lblJunk(4).Caption = "This program is Open Source and free to use, modify or redistribute under the GNU General Public License. Any violation to the license terms might result in severe civil and criminal penalties."
End Sub

Private Sub lblJunk_Click(Index As Integer)
    On Error Resume Next
    Unload Me
End Sub
