VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   4  'Icon
   ScaleHeight     =   6555
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdata 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.CommandButton cmd_add_new 
      Caption         =   "&Add New"
      Height          =   615
      Left            =   10440
      TabIndex        =   5
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdpendingwork 
      Caption         =   "Pending Work"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lbl_delete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lbl_delete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lbl_edit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   2
      Left            =   0
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbl_edit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   1
      Left            =   0
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbl_content 
      Caption         =   "012345678901234567890123456789012345678901234567890123456789"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   11700
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_content 
      Caption         =   "012345678901234567890123456789012345678901234567890123456789"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   11700
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_no_work 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Pending Work"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1050
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   10995
   End
   Begin VB.Label lbl_content 
      Caption         =   "012345678901234567890123456789012345678901234567890123456789"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   11700
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_delete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   0
      Left            =   11880
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lbl_edit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   0
      Left            =   11400
      MouseIcon       =   "Form1.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total_work As Integer

Private Sub Form_Load()
    Dim i As Integer
    lbl_height = lbl_content(0).Height
    If GetSetting("remindME", "Pending Work", "Total Work", "0") = "0" Then SaveSetting "remindME", "Pending Work", "Total Work", "0"
    total_work = Val(GetSetting("remindME", "Pending Work", "Total Work"))
    If total_work = 0 Then
        lbl_no_work.Visible = True
    ElseIf total_work <= 3 Then
        For i = 1 To total_work
            lbl_content(i - 1).Visible = True
            lbl_delete(i - 1).Visible = True
            lbl_edit(i - 1).Visible = True
            lbl_delete(i - 1).Move lbl_content(i - 1).Left + lbl_content(i - 1).Width - lbl_delete(i - 1).Width, lbl_content(i - 1).Top + lbl_content(i - 1).Height + 10
            lbl_edit(i - 1).Move lbl_delete(i - 1).Left - lbl_edit(i - 1).Width - 100, lbl_delete(i - 1).Top
            If i < 3 Then
                lbl_content(i).Move lbl_content(i - 1).Left, lbl_delete(i - 1).Top + lbl_delete(i - 1).Height + 30
            End If
        Next i
    End If
End Sub

Private Sub lbl_edit_Click(Index As Integer)
txtdata.Move lbl_content(Index).Left, lbl_content(Index).Top, lbl_content(Index).Width, lbl_content(Index).Height
txtdata.Visible = True
txtdata.Text = lbl_content(Index).Caption
lbl_edit(Index).Caption = "Done Editing"
End Sub

