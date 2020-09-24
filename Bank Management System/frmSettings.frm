VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   120
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7095
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Bank Settings"
      TabPicture(0)   =   "frmSettings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(6)"
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(2)=   "Text1(5)"
      Tab(0).Control(3)=   "Text1(4)"
      Tab(0).Control(4)=   "Text1(3)"
      Tab(0).Control(5)=   "Text1(2)"
      Tab(0).Control(6)=   "Text1(1)"
      Tab(0).Control(7)=   "Text1(0)"
      Tab(0).Control(8)=   "dialog"
      Tab(0).Control(9)=   "lblFieldLabel(8)"
      Tab(0).Control(10)=   "lblFieldLabel(7)"
      Tab(0).Control(11)=   "lblFieldLabel(5)"
      Tab(0).Control(12)=   "lblFieldLabel(4)"
      Tab(0).Control(13)=   "lblFieldLabel(3)"
      Tab(0).Control(14)=   "lblFieldLabel(2)"
      Tab(0).Control(15)=   "lblFieldLabel(1)"
      Tab(0).Control(16)=   "lblFieldLabel(0)"
      Tab(0).Control(17)=   "lblFieldLabel(6)"
      Tab(0).Control(18)=   "picLogo"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Account Settings"
      TabPicture(1)   =   "frmSettings.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblFieldLabel(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblFieldLabel(10)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblFieldLabel(11)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblFieldLabel(12)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblFieldLabel(13)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdNavigate(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdOptions(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1(7)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text1(8)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text1(9)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text1(10)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text1(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   11
         Left            =   2640
         TabIndex        =   30
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   10
         Left            =   2640
         TabIndex        =   29
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   765
         Index           =   9
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   8
         Left            =   2640
         TabIndex        =   27
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   7
         Left            =   2640
         TabIndex        =   26
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   6
         Left            =   -73320
         TabIndex        =   20
         Top             =   3120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -73320
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   5
         Left            =   -73320
         TabIndex        =   17
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   4
         Left            =   -73320
         TabIndex        =   16
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   3
         Left            =   -73320
         TabIndex        =   15
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   2
         Left            =   -73320
         TabIndex        =   14
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   1
         Left            =   -73320
         TabIndex        =   13
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   0
         Left            =   -73320
         TabIndex        =   12
         Top             =   600
         Width           =   2895
      End
      Begin MSComDlg.CommonDialog dialog 
         Left            =   -71640
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdOptions 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   0
         Left            =   360
         Picture         =   "frmSettings.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavigate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   0
         Left            =   2160
         Picture         =   "frmSettings.frx":6B72
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3720
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Balance :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   13
         Left            =   900
         TabIndex        =   25
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Rate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   12
         Left            =   1290
         TabIndex        =   24
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   11
         Left            =   1410
         TabIndex        =   23
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   10
         Left            =   1170
         TabIndex        =   22
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Id :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   9
         Left            =   1485
         TabIndex        =   21
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   8
         Left            =   -74355
         TabIndex        =   19
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click the Image to Change it"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   450
         Index           =   7
         Left            =   -70320
         TabIndex        =   11
         Top             =   2880
         Width           =   2205
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   5
         Left            =   -74025
         TabIndex        =   10
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address 3 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   4
         Left            =   -74460
         TabIndex        =   9
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address2 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   3
         Left            =   -74415
         TabIndex        =   8
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   2
         Left            =   -74460
         TabIndex        =   7
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Number :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   1
         Left            =   -74880
         TabIndex        =   6
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   0
         Left            =   -74640
         TabIndex        =   5
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   210
         Index           =   6
         Left            =   -74265
         TabIndex        =   4
         Top             =   2760
         Width           =   750
      End
      Begin VB.Image picLogo 
         Height          =   2295
         Left            =   -70320
         Picture         =   "frmSettings.frx":CE5D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Settings"
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
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmSettings
Public rsSettings As Recordset

Private Sub cmdNavigate_Click(Index As Integer)
    Select Case Index
        Case Is = 0 'Move First
            Call MoveToFirst(rsAccTypes)
            Call DisplayCustomers(rsAccTypes)
        Case Is = 1 'Move Previous
            Call MoveToPrev(rsAccTypes)
            Call DisplayCustomers(rsAccTypes)
        Case Is = 2 'Move Next
            Call MoveToNext(rsAccTypes)
            Call DisplayCustomers(rsAccTypes)
        Case Is = 3 'Move Last
            Call MoveToLast(rsAccTypes)
            Call DisplayCustomers(rsAccTypes)
    End Select
End Sub

Public Sub DisplayCustomers(myrs As Recordset)
    For i = 7 To 11: Text1(i).Text = myrs(i - 7): Next i
End Sub

Private Sub cmdOptions_Click(Index As Integer)
    Select Case Index
        Case Is = 0 'Add
            For i = 0 To 4: cmdOptions(i).Enabled = False: Next i
            cmdOptions(1).Enabled = True: cmdOptions(2).Enabled = True
            Call UnLock_Form_Controls(Me)
            Call clear_Form_Controls(Me)
            rsAccTypes.AddNew
        Case Is = 1 'Save
            For i = 7 To 11
                If Text1(i).Text = "" Then Call Messager
            Next i
            
            For i = 7 To 11
                rsAccTypes(i - 7) = Text1(i).Text
            Next i
            
            For i = 0 To 4: cmdOptions(i).Enabled = True: Next i
            cmdOptions(1).Enabled = False: cmdOptions(2).Enabled = False
            Call Lock_Form_Controls(Me)
            
        Case Is = 2 'Cancel
            rsAccTypes.CancelUpdate
            For i = 0 To 4: cmdOptions(i).Enabled = True: Next i
            cmdOptions(1).Enabled = False: cmdOptions(2).Enabled = False
            
            Call DisplayCustomers(rsAccTypes)
            Call Lock_Form_Controls(Me)
            
        Case Is = 3 'Edit
            For i = 0 To 4: cmdOptions(i).Enabled = False: Next i
            cmdOptions(1).Enabled = True: cmdOptions(2).Enabled = True
            Call UnLock_Form_Controls(Me)
            
        Case Is = 4 'Delete
            If (MsgBox("Sure To Delete?", vbYesNo + vbQuestion, "Confirm Delete")) = vbYes Then
                With rsAccTypes
                    If .BOF = True And .EOF = True Then
                        MsgBox "Nothing  to delete", vbInformation
                        Exit Sub
                    End If
                    .Delete
                    Call clear_Form_Controls(Me)
                    .MoveFirst
                    Call DisplayCustomers(rsAccTypes)
                End With
                Call Lock_Form_Controls(Me)
            End If
    End Select
End Sub

Private Sub Form_Load()
    Call create_navigation_buttons(Me)
    picLogo.Picture = LoadPicture(App.Path & "\pictures\logo.jpg")
    Call create_countries(Combo1)
    Call connectDatabase
    
    Set rsSettings = New ADODB.Recordset
    rsSettings.Open "tblSettings", cnBank, adOpenKeyset, adLockOptimistic
    Call Lock_Form_Controls(Me)
    Call DisplayCustomers(rsAccTypes)
End Sub

Private Sub picLogo_DblClick()
    With dialog
        .DialogTitle = "Choose the Picture"
        .FileName = ""
        On Error GoTo e
        .ShowOpen
        If .FileName = "" Then Exit Sub
    End With
    
    picLogo.Picture = LoadPicture(dialog.FileName)
    FileCopy dialog.FileName, App.Path & "\pictures\logo.jpg"
e:
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 1 Then
        For i = 0 To 6
            If i < 6 Then Text1(i).Text = rsSettings(i)
            If i = 6 Then
                Combo1.Text = rsSettings(i)
                Text1(i).Text = rsSettings(i + 1)
            End If
        Next i
        Call UnLock_Form_Controls(Me)
    Else
        For i = 0 To 6
            If i < 6 Then rsSettings(i) = Text1(i).Text
            If i = 6 Then
                rsSettings(i) = Combo1.Text
                rsSettings(i + 1) = Text1(i).Text
            End If
        Next i
        rsSettings.Update
    End If
End Sub

