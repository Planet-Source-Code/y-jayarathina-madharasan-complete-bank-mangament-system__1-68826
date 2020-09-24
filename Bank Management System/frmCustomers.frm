VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCustomers.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   7050
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   32
      Top             =   6585
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboAccType 
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox cboContactTitle 
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   10
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2280
      Width           =   2040
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   9
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4080
      Width           =   2400
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   4080
      Width           =   2280
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   7
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3240
      Width           =   1440
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   6
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3240
      Width           =   1800
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3240
      Width           =   2280
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2280
      Width           =   1440
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   3
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   2280
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   840
      Width           =   2520
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   840
      Width           =   2280
   End
   Begin VB.CommandButton cmdNavigate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   2175
      Picture         =   "frmCustomers.frx":B2B6
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   0
      Left            =   360
      Picture         =   "frmCustomers.frx":115A1
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   0
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker txtDateJoined 
      Height          =   255
      Left            =   4485
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   71041027
      CurrentDate     =   39178
   End
   Begin VB.Label lblMin 
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   10
      Left            =   105
      TabIndex        =   17
      Top             =   3840
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   11
      Left            =   4440
      TabIndex        =   16
      Top             =   3840
      Width           =   1245
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location / Town:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   8
      Left            =   4440
      TabIndex        =   15
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PostalCode:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   14
      Top             =   3000
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   5
      Left            =   2520
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   13
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   14
      Left            =   4440
      TabIndex        =   10
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   4500
      TabIndex        =   8
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   2070
      TabIndex        =   7
      Top             =   600
      Width           =   930
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "National ID NO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   4
      Left            =   2085
      TabIndex        =   5
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ContactTitle:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   165
      TabIndex        =   4
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DateJoined:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   12
      Left            =   4485
      TabIndex        =   3
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customers"
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
      Width           =   7095
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmCustomers
Private Sub cboAccType_Click()
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select * FROM tblAccTypes WHERE AccountName='" & cboAccType.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
If rsTemp.RecordCount > 0 Then lblMin = rsTemp(4)
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
    Select Case Index
        Case Is = 0 'Move First
            Call MoveToFirst(rsCustomers)
            Call DisplayCustomers(rsCustomers)
        Case Is = 1 'Move Previous
            Call MoveToPrev(rsCustomers)
            Call DisplayCustomers(rsCustomers)
        Case Is = 2 'Move Next
            Call MoveToNext(rsCustomers)
            Call DisplayCustomers(rsCustomers)
        Case Is = 3 'Move Last
            Call MoveToLast(rsCustomers)
            Call DisplayCustomers(rsCustomers)
    End Select
End Sub

Private Sub cmdOptions_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    Case Is = 0 'Add
        For i = 0 To 4: cmdOptions(i).Enabled = False: Next i
        cmdOptions(1).Enabled = True: cmdOptions(2).Enabled = True
        Call UnLock_Form_Controls(Me)
        Call clear_Form_Controls(Me)
        Call GenerateNewCustomerCode
        txtCustomer(0).Locked = True
        txtCustomer(1).SetFocus
        rsCustomers.AddNew
        txtDateJoined.Value = Date
    Case Is = 1 'Save
        If txtCustomer(9).Text = "" Then txtCustomer(9).Text = "N/A"
        For i = 0 To 10
            If txtCustomer(i).Text = "" Then
                Call Messager
                txtCustomer(i).SetFocus
                Exit Sub
            End If
        Next i
                
        If CCur(txtCustomer(10).Text) < CCur(lblMin.Caption) Then
            MsgBox "Opening balance should be atleast " & lblMin.Caption & " for this type of Account", vbInformation
            Exit Sub
        End If
        z = 0
        For X = 0 To 13
            Select Case X
                Case Is = 3: rsCustomers(X) = cboContactTitle.Text
                Case Is = 11: rsCustomers(X) = txtDateJoined.Value
                Case Is = 12: rsCustomers(X) = cboAccType.Text
                Case Else:  rsCustomers(X) = txtCustomer(z).Text: z = z + 1
            End Select
        Next X
        For i = 0 To 4: cmdOptions(i).Enabled = True: Next i
        cmdOptions(1).Enabled = False: cmdOptions(2).Enabled = False
        Call Lock_Form_Controls(Me)
        
    Case Is = 2 'Cancel
        rsCustomers.CancelUpdate
        For i = 0 To 4: cmdOptions(i).Enabled = True: Next i
        cmdOptions(1).Enabled = False: cmdOptions(2).Enabled = False
        
        Call DisplayCustomers(rsCustomers)
        Call Lock_Form_Controls(Me)
    Case Is = 3 'Edit
        For i = 0 To 4: cmdOptions(i).Enabled = False: Next i
        cmdOptions(1).Enabled = True: cmdOptions(2).Enabled = True
        Call UnLock_Form_Controls(Me)
        txtCustomer(0).Locked = True
    Case Is = 4 'Delete
        If (MsgBox("Sure To Delete?", vbYesNo + vbQuestion, "Confirm Delete")) = vbYes Then
            With rsCustomers
                If .BOF = True And .EOF = True Then
                    MsgBox "Nothing  to delete", vbInformation
                    Exit Sub
                End If
                .Delete
                Call clear_Form_Controls(Me)
                .MoveFirst
                Call DisplayCustomers(rsCustomers)
            End With
            Call Lock_Form_Controls(Me)
        End If
    End Select
    End Sub
    Public Sub GenerateNewCustomerCode()
    Dim lastnumber As Long, newnumber As Long
    With rsCustomers
    If .BOF = True And .EOF = True Then
        lastnumber = 2004000
    Else
        .MoveLast
        lastnumber = !CustomerID
    End If
    
    newnumber = lastnumber + 1
    txtCustomer(0).Text = newnumber
    End With
End Sub
Private Sub Form_Load()
    Call create_navigation_buttons(Me)
    txtDateJoined.Value = Date
    With cboContactTitle
        .AddItem "MR."
        .AddItem "MRS."
        .AddItem "MISS."
        .AddItem "DR."
        .AddItem "PROFF."
        .AddItem "SIR."
        .AddItem "REV."
        .AddItem "FR."
    End With
    Call connectDatabase
    For X = 1 To rsAccTypes.RecordCount
        cboAccType.AddItem rsAccTypes(1) '!AccountName
        rsAccTypes.MoveNext
    Next X
    Call DisplayCustomers(rsCustomers)
    Call Lock_Form_Controls(Me)
End Sub

Public Sub DisplayCustomers(myrs As Recordset)
    Dim z As Integer
    With myrs
        If .BOF = True And .EOF = True Then Exit Sub
        On Error Resume Next
        z = 0
        For X = 0 To 13
            Select Case X
                Case Is = 3: cboContactTitle.Text = myrs(X)
                Case Is = 11: txtDateJoined.Value = myrs(X)
                Case Is = 12: cboAccType.Text = myrs(X)
                Case Else: txtCustomer(z).Text = myrs(X): z = z + 1
            End Select
        Next X
        StatusBar1.SimpleText = CStr("Record :" & .AbsolutePosition & " of " & .RecordCount)
    End With
End Sub

Private Sub txtCustomer_LostFocus(Index As Integer)
    If Index = 10 Then
        If IsNumeric(txtCustomer(10).Text) = False And Not (txtCustomer(10).Text = "") Then
            MsgBox "Invalid Input", vbOKOnly + vbCritical, "Error"
            txtCustomer(10).Text = ""
            txtCustomer(10).SetFocus
        End If
    End If
End Sub
