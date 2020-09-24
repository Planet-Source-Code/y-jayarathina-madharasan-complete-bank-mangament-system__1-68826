VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeposits 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposits"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmDeposits.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   6855
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Print Bill of this Transaction"
      Height          =   615
      Left            =   4440
      TabIndex        =   23
      Top             =   3840
      Width           =   2295
   End
   Begin VB.OptionButton optCash 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Others Specify..."
      Height          =   315
      Index           =   2
      Left            =   4080
      TabIndex        =   21
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton optCash 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cheque"
      Height          =   315
      Index           =   1
      Left            =   2640
      TabIndex        =   20
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   19
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   17
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   885
      Index           =   2
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.OptionButton optCash 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cash"
      Height          =   315
      Index           =   0
      Left            =   1590
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   71041027
      CurrentDate     =   38293
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposits"
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
      TabIndex        =   13
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   12
      Top             =   480
      Width           =   960
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmountDeposited:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   5
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label lblBalance 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Balance in this Account :"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "frmDeposits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmDeposits
Dim currBalance As Currency
Private Sub cboCustomerNo_Click()
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select * FROM tblCustomers WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
    With rsTemp
        If .RecordCount > 0 Then
            txtDeposits(1).Text = !AccountNo
            txtDeposits(2).SetFocus
        Else
            MsgBox "Invalid Customer Code", vbInformation
            txtDeposits(1).Text = ""
        Exit Sub
        End If
        .Close
    End With
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select * FROM tblCustomers WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
    If rsTemp.RecordCount > 0 Then
        lblBalance.Caption = rsTemp(13)
    Else
        Exit Sub
    End If
    rsTemp.Close
End Sub


Private Sub cmdOptions_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    Case Is = 0
        If cboCustomerNo.Text = "" Then
            Call Messager
            cboCustomerNo.SetFocus
            Exit Sub
        End If
        
        If txtDeposits(1).Text = "" Then
            MsgBox "Please Enter the Customer Correctly.", vbInformation
            cboCustomerNo.SetFocus
            Exit Sub
        End If
        
        For i = 0 To 5
            If txtDeposits(i).Text = "" Then
                Call Messager
                txtDeposits(i).SetFocus
                Exit Sub
            End If
        Next i
        
        i = 0
        For X = 0 To 7
            Select Case X
            Case Is = 1
                rsDeposit(X) = cboCustomerNo.Text
            Case Is = 7
                rsDeposit(X) = txtDated.Value
            Case Else
                rsDeposit(X) = txtDeposits(i).Text
                i = i + 1
            End Select
        Next X
        rsDeposit.Update
        currBalance = (CCur(lblBalance.Caption) + CCur(txtDeposits(3).Text))
    
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select * FROM tblCustomers WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
        If rsTemp.RecordCount > 0 Then
             rsTemp(13) = currBalance
        Else
            Exit Sub
        End If
         rsTemp.Update
         rsTemp.Close
         
        If Check1.Value = 1 Then
            DataEnvironment1.rsCommand1.Filter = "TransactionID='" & txtDeposits(0).Text & "'"
            DataReport1.Show
        End If
    Case Is = 1
        rsDeposit.CancelUpdate
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    With cmdOptions(0)
        .Picture = LoadPicture(App.Path & "\pictures\n2.jpg")
    End With
    For i = 1 To 1
        Load cmdOptions(i)
        With cmdOptions(i)
            .Visible = True
            .Left = cmdOptions(i - 1).Left + cmdOptions(i - 1).Width + 25
            If i = 1 Then
                .Picture = LoadPicture(App.Path & "\pictures\n3.jpg")
            Else
                .Picture = LoadPicture(App.Path & "\pictures\n10.jpg")
            End If
        End With
    Next i
    
    Call connectDatabase
    
    For X = 1 To rsCustomers.RecordCount
        cboCustomerNo.AddItem rsCustomers(0)
        rsCustomers.MoveNext
    Next X
    Call clear_Form_Controls(Me)
    Call GenerateNewTransactCode
    rsDeposit.AddNew
End Sub

Public Sub GenerateNewTransactCode()
    Dim lastnumber As Long, newnumber As Long
    If rsDeposit.BOF = True And rsDeposit.EOF = True Then
        lastnumber = 1000
    Else
        rsDeposit.MoveLast
        lastnumber = rsDeposit(0)
    End If
    newnumber = lastnumber + 1
    txtDeposits(0).Text = newnumber
End Sub

Private Sub optCash_Click(Index As Integer)
    Select Case Index
    Case Is = 0
        txtDeposits(5).Text = "N/A"
        txtDeposits(5).Locked = True
        txtDeposits(4).Text = "CASH"
        txtDeposits(4).Locked = True
    Case Is = 1
        txtDeposits(4).Text = "CHEQUE"
        txtDeposits(4).Locked = True
        txtDeposits(5).Text = ""
        txtDeposits(5).Locked = False
        txtDeposits(5).SetFocus
    Case Is = 2
        txtDeposits(4).Locked = False
        txtDeposits(4).SetFocus
        txtDeposits(4).Text = ""
        txtDeposits(5).Text = "N/A"
        txtDeposits(5).Locked = True
    End Select
End Sub

Private Sub txtDeposits_LostFocus(Index As Integer)
    If Index = 3 Then
        If IsNumeric(txtDeposits(3).Text) = False And Not (txtDeposits(3).Text = "") Then
            MsgBox "Invalid Input", vbOKOnly + vbCritical, "Error"
            txtDeposits(3).Text = ""
            txtDeposits(3).SetFocus
        End If
    End If
End Sub
