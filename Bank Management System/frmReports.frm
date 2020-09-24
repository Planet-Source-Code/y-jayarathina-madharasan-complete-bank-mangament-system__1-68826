VERSION 5.00
Begin VB.Form frmReports 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3705
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transactions Report"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Account Types Report"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contacts Report"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   4
      Left            =   -120
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   3
      Left            =   -120
      TabIndex        =   7
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose The type of report you want to Generate:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmReports

Private Sub Command1_Click()
    For i = 0 To 2
        If Option1(i).Value = True Then Exit For
    Next i
    Select Case i
        Case Is = 0
                DataReport3.Show
        Case Is = 1
                DataReport4.Show
        Case Is = 2
                frmTransactions.Show
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    Command1.Picture = LoadPicture(App.Path & "\pictures\n10.jpg")
End Sub
