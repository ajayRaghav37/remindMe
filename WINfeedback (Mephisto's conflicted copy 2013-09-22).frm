VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form WINfeedback 
   Appearance      =   0  'Flat
   BackColor       =   &H0026A5E1&
   BorderStyle     =   0  'None
   Caption         =   "ANIco.in Feedback Management Sytem 1.0"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H0026A5E1&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog DlgBox 
      Left            =   360
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Add attachment(s)"
   End
   Begin VB.Frame frmJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      ForeColor       =   &H80000008&
      Height          =   840
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   4050
      Width           =   840
      Begin VB.Label cmdRemoveAll 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Remove All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00202020&
         Height          =   180
         Left            =   90
         TabIndex        =   41
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblAttachCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00202020&
         Height          =   750
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   40
         Top             =   90
         Width           =   840
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3600
   End
   Begin VB.Frame frmJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   2790
      Width           =   3600
      Begin VB.Label cmdRating 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   15
         TabIndex        =   10
         Top             =   1275
         Width           =   3585
      End
      Begin VB.Label cmdRating 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   15
         TabIndex        =   11
         Top             =   960
         Width           =   3585
      End
      Begin VB.Label cmdRating 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   15
         TabIndex        =   12
         Top             =   645
         Width           =   3585
      End
      Begin VB.Label cmdRating 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   15
         TabIndex        =   13
         Top             =   330
         Width           =   3585
      End
      Begin VB.Label cmdRating 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   15
         TabIndex        =   14
         Top             =   15
         Width           =   3585
      End
      Begin VB.Label lblRatingType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Overall"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblRatingType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ease of Use"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1005
         Width           =   885
      End
      Begin VB.Label lblRatingType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "System Load"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblRatingType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Interface"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   375
         Width           =   1080
      End
      Begin VB.Label lblRatingType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Functionality"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   60
         Width           =   1035
      End
      Begin VB.Shape ShpJunk 
         BorderColor     =   &H80000006&
         Height          =   1605
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   3600
      End
      Begin VB.Label lblJunk 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   60
         Width           =   15
      End
      Begin VB.Label lblJunk 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   15
      End
      Begin VB.Label lblJunk 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   1005
         Width           =   15
      End
      Begin VB.Label lblJunk 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   690
         Width           =   15
      End
      Begin VB.Label lblJunk 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   375
         Width           =   15
      End
      Begin VB.Label lblRating 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   3420
         TabIndex        =   19
         Top             =   75
         Width           =   45
      End
      Begin VB.Label lblRating 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   3420
         TabIndex        =   18
         Top             =   390
         Width           =   45
      End
      Begin VB.Label lblRating 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   3420
         TabIndex        =   17
         Top             =   705
         Width           =   45
      End
      Begin VB.Label lblRating 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   3420
         TabIndex        =   16
         Top             =   1020
         Width           =   45
      End
      Begin VB.Label lblRating 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Index           =   4
         Left            =   3420
         TabIndex        =   15
         Top             =   1335
         Width           =   45
      End
      Begin VB.Shape ShpRating 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0036B5F1&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   0
         Left            =   15
         Top             =   15
         Width           =   3585
      End
      Begin VB.Shape ShpRating 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0046C5FF&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   1
         Left            =   15
         Top             =   330
         Width           =   3585
      End
      Begin VB.Shape ShpRating 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0036B5F1&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   2
         Left            =   15
         Top             =   645
         Width           =   3585
      End
      Begin VB.Shape ShpRating 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0046C5FF&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   3
         Left            =   15
         Top             =   960
         Width           =   3585
      End
      Begin VB.Shape ShpRating 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0036B5F1&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   4
         Left            =   15
         Top             =   1275
         Width           =   3585
      End
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "[Fill in if you want a reply]"
      Top             =   600
      Width           =   3600
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   1605
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   3600
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "&Criticism"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   39
      Top             =   2460
      Width           =   960
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "&Praise"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   38
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "&New feature"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   1950
      Width           =   960
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bug report"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   1695
      Width           =   960
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "&Miscellaneous"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label cmdSend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sen&d"
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
      Left            =   2325
      TabIndex        =   7
      Tag             =   "o"
      Top             =   4500
      Width           =   1125
   End
   Begin VB.Label cmdHide 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0026A5E1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Hide"
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
      Left            =   3555
      TabIndex        =   8
      Tag             =   "o"
      Top             =   4500
      Width           =   1125
   End
   Begin VB.Label cmdAttach 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Attach File(s)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   3870
      Width           =   765
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
      Height          =   5010
      Index           =   10
      Left            =   0
      TabIndex        =   30
      Tag             =   "o"
      Top             =   0
      Width           =   4800
   End
   Begin VB.Shape ShpJunk 
      BorderColor     =   &H80000006&
      Height          =   390
      Index           =   1
      Left            =   1080
      Top             =   4500
      Width           =   1125
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0026A5E1&
      BackStyle       =   0  'Transparent
      Caption         =   "FMS BETA 1"
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
      Left            =   1335
      TabIndex        =   6
      Top             =   4710
      Width           =   615
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0026A5E1&
      BackStyle       =   0  'Transparent
      Caption         =   "ANIco.in"
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
      Left            =   1230
      TabIndex        =   5
      Top             =   4515
      Width           =   825
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ratings"
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
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      UseMnemonic     =   0   'False
      Width           =   645
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
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
      Left            =   120
      TabIndex        =   3
      Top             =   630
      UseMnemonic     =   0   'False
      Width           =   705
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Index           =   9
      Left            =   120
      TabIndex        =   31
      Top             =   150
      UseMnemonic     =   0   'False
      Width           =   525
   End
   Begin VB.Label lblJunk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
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
      Index           =   11
      Left            =   120
      TabIndex        =   34
      Top             =   1110
      UseMnemonic     =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "WINfeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TypeSel As Byte
Dim RateVal(4) As Single
Dim Comment(4) As String
Dim Attachments As String
Dim ValidMail As Boolean
Const MaxRateWidth As Single = 3585
Const Part50th As Single = MaxRateWidth / 50
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Boolean
Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Sub lblAttachCount_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TempNum As Integer
    For TempNum = 1 To Data.Files.Count
        Attachments = Attachments & Data.Files(TempNum) & "*"
        lblAttachCount.Caption = Trim$(Val(lblAttachCount.Caption) + 1)
    Next
End Sub

Private Sub cmdAttach_Click()
    Dim SelFiles() As String
    Dim TempNum As Integer
    DlgBox.Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
    DlgBox.ShowOpen
    If Len(DlgBox.FileName) <> 0 Then
        SelFiles() = Split(DlgBox.FileName, vbNullChar)
        If UBound(SelFiles) = 0 Then
            Attachments = Attachments & DlgBox.FileName & "*"
            lblAttachCount.Caption = Trim$(Val(lblAttachCount.Caption) + 1)
        Else
            For TempNum = 1 To UBound(SelFiles)
                Attachments = Attachments & SelFiles(0) & "\" & SelFiles(TempNum) & "*"
            Next
            lblAttachCount.Caption = Trim$(Val(lblAttachCount.Caption) + UBound(SelFiles))
        End If
        cmdRemoveAll.Enabled = True
    End If
End Sub

Private Sub cmdHide_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Attachments = "*" & Attachments
    Enabled = False
    txtDesc.Enabled = False
    txtName.Enabled = False
    txtMail.Enabled = False
    cmdSend.Caption = "Wait..."
    SendFeedback
End Sub

Private Sub cmdSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cBrighter Then
        cmdSend.BackColor = cBrighter
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub cmdHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdHide.BackColor <> cBrighter Then
        cmdHide.BackColor = cBrighter
    End If
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
End Sub

Private Sub cmdHide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub cmdSend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
End Sub

Private Sub cmdRating_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        UpdateRate Index, X
    End If
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub cmdRating_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        UpdateRate Index, X
    End If
End Sub

Private Sub cmdRemoveAll_Click()
    Attachments = vbNullString
    lblAttachCount.Caption = "+"
    cmdRemoveAll.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyF
                If RateVal(0) < 5 Then
                    UpdateRate 0, ShpRating(0).Width + Part50th
                Else
                    UpdateRate 0, 0
                End If
            Case vbKeyU
                If RateVal(1) < 5 Then
                    UpdateRate 1, ShpRating(1).Width + Part50th
                Else
                    UpdateRate 1, 0
                End If
            Case vbKeyS
                If RateVal(2) < 5 Then
                    UpdateRate 2, ShpRating(2).Width + Part50th
                Else
                    UpdateRate 2, 0
                End If
            Case vbKeyE
                If RateVal(3) < 5 Then
                    UpdateRate 3, ShpRating(3).Width + Part50th
                Else
                    UpdateRate 3, 0
                End If
            Case vbKeyO
                If RateVal(4) < 5 Then
                    UpdateRate 4, ShpRating(4).Width + Part50th
                Else
                    UpdateRate 4, 0
                End If
        End Select
    ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE
                If RateVal(3) > 0 Then
                    UpdateRate 3, ShpRating(3).Width - Part50th
                Else
                    UpdateRate 3, MaxRateWidth
                End If
            Case vbKeyF
                If RateVal(0) > 0 Then
                    UpdateRate 0, ShpRating(0).Width - Part50th
                Else
                    UpdateRate 0, MaxRateWidth
                End If
            Case vbKeyO
                If RateVal(4) > 0 Then
                    UpdateRate 4, ShpRating(4).Width - Part50th
                Else
                    UpdateRate 4, MaxRateWidth
                End If
            Case vbKeyS
                If RateVal(2) > 0 Then
                    UpdateRate 2, ShpRating(2).Width - Part50th
                Else
                    UpdateRate 2, MaxRateWidth
                End If
            Case vbKeyU
                If RateVal(1) > 0 Then
                    UpdateRate 1, ShpRating(1).Width - Part50th
                Else
                    UpdateRate 1, MaxRateWidth
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyA
                cmdAttach_Click
            Case vbKeyD
                cmdSend_Click
            Case vbKeyH
                cmdHide_Click
            Case vbKeyR
                cmdRemoveAll_Click
            
            Case vbKeyM
                lblType_Click (0)
            Case vbKeyB
                lblType_Click (1)
            Case vbKeyN
                lblType_Click (2)
            Case vbKeyP
                lblType_Click (3)
            Case vbKeyC
                lblType_Click (4)
        End Select
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Top = WINremindME.Top + WINremindME.lblFeedback.Top - Height
    Left = WINremindME.Left + WINremindME.lblFeedback.Left - (Width - WINremindME.lblFeedback.Width) / 2
    
    'Checking internet connectivity
    
    If WinSock1.LocalIP = "127.0.0.1" Or WinSock1.LocalIP = vbNullString Then
        Dim objControl As Control
        For Each objControl In Me.Controls
            If objControl.Name <> "cmdHide" Then
                objControl.Enabled = False
            End If
        Next
        txtDesc.Text = "NO INTERNET ACCESS"
    End If
    
    txtName.Text = Environ$("USERNAME")
    UpdateRate 0, 39 * Part50th
    UpdateRate 1, 38 * Part50th
    UpdateRate 2, 49 * Part50th
    UpdateRate 3, 33 * Part50th
    UpdateRate 4, (ShpRating(0).Width + ShpRating(1).Width + ShpRating(2).Width + ShpRating(3).Width) / 4
End Sub

Private Sub frmJunk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub lblAttachCount_Click()
    cmdAttach_Click
End Sub

Private Sub lblJunk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub lblType_Click(Index As Integer)
    If Index = TypeSel Then
        Exit Sub
    End If
    lblType(TypeSel).FontBold = False
    Comment(TypeSel) = txtDesc.Text
    TypeSel = Index
    lblType(TypeSel).FontBold = True
    txtDesc.Text = Comment(TypeSel)
End Sub

Private Sub UpdateRate(RateIndex As Integer, RateWidth As Single)
    Dim RateStr As String
    If RateWidth <= MaxRateWidth Then
        If RateWidth >= 15 Then
            ShpRating(RateIndex).Width = RateWidth
            RateVal(RateIndex) = Round(5 * RateWidth / MaxRateWidth, 1)
        Else
            ShpRating(RateIndex).Width = 15
            RateVal(RateIndex) = 0
        End If
    Else
            ShpRating(RateIndex).Width = MaxRateWidth
            RateVal(RateIndex) = 5
    End If
    Select Case RateIndex
        Case 0
            Select Case RateVal(RateIndex)
                Case Is <= 1
                    RateStr = "Poor ("
                Case Is <= 2
                    RateStr = "Less useful ("
                Case Is <= 3
                    RateStr = "Satisfactory ("
                Case Is <= 4
                    RateStr = "Professional ("
                Case Else
                    RateStr = "Flawless ("
            End Select
        Case 1
            Select Case RateVal(RateIndex)
                Case Is <= 1
                    RateStr = "Ewww ("
                Case Is <= 2
                    RateStr = "Dislike ("
                Case Is <= 3
                    RateStr = "Just good ("
                Case Is <= 4
                    RateStr = "Nice ("
                Case Else
                    RateStr = "Awesome ("
            End Select
        Case 2
            Select Case RateVal(RateIndex)
                Case Is <= 1
                    RateStr = "Very heavy ("
                Case Is <= 2
                    RateStr = "Heavy ("
                Case Is <= 3
                    RateStr = "Moderate ("
                Case Is <= 4
                    RateStr = "Low ("
                Case Else
                    RateStr = "Ultra-Light ("
            End Select
        Case 3
            Select Case RateVal(RateIndex)
                Case Is <= 1
                    RateStr = "Impossible ("
                Case Is <= 2
                    RateStr = "Difficult ("
                Case Is <= 3
                    RateStr = "Easy ("
                Case Is <= 4
                    RateStr = "Quite easy ("
                Case Else
                    RateStr = "Kid's stuff ("
            End Select
        Case 4
            Select Case RateVal(RateIndex)
                Case Is <= 1
                    RateStr = "Useless ("
                Case Is <= 2
                    RateStr = "Immature ("
                Case Is <= 3
                    RateStr = "Good enough ("
                Case Is <= 4
                    RateStr = "Useful ("
                Case Else
                    RateStr = "Must have ("
            End Select
    End Select
    lblRating(RateIndex).Caption = RateStr & Format$(RateVal(RateIndex), "0.0") & ")"
    If RateIndex <> 4 Then
        UpdateRate 4, (ShpRating(0).Width + ShpRating(1).Width + ShpRating(2).Width + ShpRating(3).Width) / 4
    End If
End Sub

Private Sub lblType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub txtDesc_Change()
    Comment(TypeSel) = txtDesc.Text
End Sub

Private Sub txtDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Sub txtMail_GotFocus()
    txtMail.FontItalic = False
    txtMail.SelStart = 0
    txtMail.SelLength = Len(txtMail.Text)
End Sub

Private Sub txtMail_LostFocus()
    txtMail.FontItalic = True
End Sub

Private Sub txtName_GotFocus()
    txtName.FontItalic = False
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_LostFocus()
    txtName.FontItalic = True
End Sub

Private Sub txtMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdSend.BackColor <> cButton_Hover Then
        cmdSend.BackColor = cButton_Hover
    End If
    If cmdHide.BackColor <> cButton_Hover Then
        cmdHide.BackColor = cButton_Hover
    End If
End Sub

Private Function CompileFeedback() As String
    CompileFeedback = "We have recieved your feedback for " & App.ProductName & ". One of our support executive will get back to you very soon." & vbNewLine & vbNewLine & "Thank you for your time. We respect your privacy. Below is the information that was sent to us. A copy of attachments (if any) is also attached in this e-mail." & vbNewLine & vbNewLine
    
    CompileFeedback = CompileFeedback & "ANIco.in Feedback Management System v1.0 Report" & vbNewLine & vbNewLine
    
    CompileFeedback = CompileFeedback & "METADATA" & vbNewLine
    CompileFeedback = CompileFeedback & "Product" & vbTab & vbTab & vbTab & "- " & App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine
    CompileFeedback = CompileFeedback & "User" & vbTab & vbTab & vbTab & "- "

    Dim ValidName As Boolean
    Dim TempNum As Byte
    ValidMail = IsValidEmail(txtMail.Text)
    ValidName = (Len(txtName.Text) <> 0)
    
    If Not ValidMail And Not ValidName Then
        CompileFeedback = CompileFeedback & "Anonymous" & vbNewLine
    Else
        If ValidMail And ValidName Then
            CompileFeedback = CompileFeedback & txtName.Text & " <" & txtMail.Text & ">" & vbNewLine
        ElseIf ValidName Then
            CompileFeedback = CompileFeedback & txtName.Text & vbNewLine
        Else
            CompileFeedback = CompileFeedback & txtMail.Text & vbNewLine
        End If
    End If
    
    If Len(Comment(0) & Comment(1) & Comment(2) & Comment(3) & Comment(4)) <> 0 Then
        CompileFeedback = CompileFeedback & vbNewLine & "USER'S COMMENTS" & vbNewLine
        For TempNum = 0 To 4
            If Len(Comment(TempNum)) <> 0 Then
                CompileFeedback = CompileFeedback & Mid$(lblType(TempNum).Caption, 2) & vbTab & IIf(TempNum = 3, vbTab, vbNullString) & vbTab & "- " & Chr$(34) & Replace$(Comment(TempNum), vbCrLf, vbCrLf & vbTab & vbTab & vbTab & "  ") & Chr$(34) & vbNewLine
            End If
        Next
    End If
    
    CompileFeedback = CompileFeedback & vbNewLine & "RATINGS" & vbNewLine
    For TempNum = 0 To 4
        CompileFeedback = CompileFeedback & lblRatingType(TempNum).Caption & vbTab & vbTab & IIf(TempNum = 4, vbTab, vbNullString) & "- " & Format$(RateVal(TempNum), "0.0") & vbNewLine
    Next
    CompileFeedback = CompileFeedback & vbNewLine
    
    CompileFeedback = CompileFeedback & "SYSTEM INFORMATION" & vbNewLine
    CompileFeedback = CompileFeedback & "Operating System" & vbTab & "- " & GetWindowsVersion & vbNewLine
    CompileFeedback = CompileFeedback & "System Type" & vbTab & vbTab & "- " & xxBit & vbNewLine
    CompileFeedback = CompileFeedback & "System Name" & vbTab & vbTab & "- " & WinSock1.LocalHostName & vbNewLine
    CompileFeedback = CompileFeedback & "System IP" & vbTab & vbTab & "- " & WinSock1.LocalIP & vbNewLine
    CompileFeedback = CompileFeedback & "Running as Admin" & vbTab & "- " & IIf(IsUserAnAdmin, "Yes", "No")
End Function

Public Function IsValidEmail(email As String) As Boolean
    Dim myAt As Integer
    Dim myAtLastPos As Integer
    Dim myDot As Integer
    Dim myDotDot As Integer
    Dim myDotAt As Integer
    Dim myAtDot As Integer
    Dim mySpace As Integer
    IsValidEmail = True
    mySpace = InStr(1, email, " ", vbTextCompare)
    myAtLastPos = InStrRev(email, "@", , vbTextCompare)
    myAt = InStr(1, email, "@", vbTextCompare)
    myAtDot = InStr(1, email, "@.", vbTextCompare)
    myDotAt = InStr(1, email, ".@", vbTextCompare)
    myDot = InStr(myAt + 2, email, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
    If myAtDot > 0 Or myDotAt > 0 Or myAtLastPos <> myAt Or mySpace > 0 Or myAt = 0 Or myDot = 0 Or myDotDot > 0 Or Right(email, 1) = "." Then IsValidEmail = False
End Function

Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case 0
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case 1
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
            Case 2
                GetWindowsVersion = "Windows NT"
                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                            Case 2
                                GetWindowsVersion = "Windows 8/Server 2012"
                        End Select
                End Select
        End Select
        GetWindowsVersion = GetWindowsVersion & " (" & osv.dwVerMajor & "." & osv.dwVerMinor & "." & osv.dwBuildNumber & ")"
    Else
        GetWindowsVersion = "Unknown"
    End If
End Function

Public Function xxBit() As String
    Dim handle As Long
    Dim is64Bit As Boolean
    is64Bit = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle <> 0 Then
        IsWow64Process GetCurrentProcess(), is64Bit
    End If
    xxBit = IIf(is64Bit, "64-Bit", "32-Bit")
End Function

Public Sub SendFeedback()
    On Error GoTo Hell
    Dim MyMail As CDO.Message
    Set MyMail = New CDO.Message
    Dim TempNum As Integer
    Dim PrevOccur As Integer
    
    MyMail.Configuration.Fields(cdoSMTPServer) = "smtp.gmail.com" '"ex7.anico.in" or "smtp.anico.in"
    MyMail.Configuration.Fields(cdoSMTPServerPort) = 25 'also tried 587 and 2525 and 995 and 465 and 993
    MyMail.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    MyMail.Configuration.Fields(cdoSendUserName) = "AmitRaghav1987@gmail.com" 'feedback@anico.in
    MyMail.Configuration.Fields(cdoSendPassword) = "michael007"
    MyMail.Configuration.Fields(cdoSMTPConnectionTimeout) = 60
    MyMail.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    MyMail.Configuration.Fields.Update

    MyMail.TextBody = CompileFeedback
    
    If ValidMail Then
        MyMail.To = txtMail.Text
        MyMail.BCC = "support@anico.in"
    Else
        MyMail.To = "support@anico.in"
    End If
    
    MyMail.From = """Amit Raghav"" <AmitRaghav1987@gmail.com>" '"""ANIco.in Feedback Management System"" <feedback@anico.in>"
    
    MyMail.Subject = "Feedback for " & App.ProductName
    
    If Len(Attachments) <> 0 Then
        For TempNum = 1 To Val(lblAttachCount.Caption)
            PrevOccur = InStr(PrevOccur + 1, Attachments, "*")
            MyMail.AddAttachment Mid$(Attachments, PrevOccur + 1, InStr(PrevOccur + 1, Attachments, "*") - PrevOccur - 1)
        Next
    End If
    
    MyMail.Send
    
    Set MyMail = Nothing
    
    cmdSend.Caption = "Sent"
    Unload Me
    Exit Sub
Hell:
    MsgBox Err.Description & vbCrLf & "Error Number: " & Err.Number
    cmdSend.Caption = "Faile&d"
    Enabled = True
    txtName.Enabled = True
    txtMail.Enabled = True
    txtDesc.Enabled = True
    DoEvents
    Dim TempTimer As Single
    TempTimer = Timer
    Do
        DoEvents
    Loop Until (Timer - TempTimer > 2)
    cmdSend.Caption = "Sen&d"
End Sub
