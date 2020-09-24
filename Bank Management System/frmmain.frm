VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Bank Management System"
   ClientHeight    =   3300
   ClientLeft      =   4500
   ClientTop       =   3720
   ClientWidth     =   5940
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   2461
      ButtonWidth     =   609
      ButtonHeight    =   2302
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Cmdmenu 
         Height          =   1215
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   5940
      TabIndex        =   1
      Top             =   1395
      Width           =   5940
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MDIForm1
Private Sub Cmdmenu_Click(Index As Integer)
Select Case Index
    Case 0: frmCustomers.Show
    Case 1: frmDeposits.Show
    Case 2: frmWithdrawal.Show
    Case 3: frmTransactions.Show
    Case 4: frmReports.Show
    Case 5: frmSettings.Show
End Select
End Sub

Private Sub MDIForm_Load()
    Cmdmenu(0).Picture = LoadPicture(App.Path & "\pictures\1.jpg")
    For i = 1 To 5
        Load Cmdmenu(i)
        Cmdmenu(i).Visible = True
        Cmdmenu(i).Left = Cmdmenu(i - 1).Left + Cmdmenu(i - 1).Width + 500
        Cmdmenu(i).Picture = LoadPicture(App.Path & "\pictures\" & (i + 1) & ".jpg")
    Next i
End Sub
