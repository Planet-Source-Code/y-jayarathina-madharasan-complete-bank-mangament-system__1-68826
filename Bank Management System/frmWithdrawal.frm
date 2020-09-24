VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWithdrawal 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Withdrawl"
   ClientHeight    =   4425
   ClientLeft      =   5295
   ClientTop       =   3060
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6810
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   885
      Index           =   2
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtDeposits 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   1
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Print Bill of this Transaction"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   71041027
      CurrentDate     =   38293
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Balance in this Account :"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblBalance 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   14
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Withdrawn:"
      Height          =   195
      Index           =   4
      Left            =   30
      TabIndex        =   10
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   9
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      Caption         =   "Withdrawal"
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
      TabIndex        =   8
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmWithdrawal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmWithdrawal
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
        
        For i = 0 To 3
            If txtDeposits(i).Text = "" Then
                Call Messager
                txtDeposits(i).SetFocus
                Exit Sub
            End If
        Next i
        If (Val(lblBalance.Caption) - Val(txtDeposits(3).Text)) < Val(Label3.Caption) Then
            MsgBox "Money cant be withdrawn as the account balance has reached minimum!"
            Exit Sub
        End If
        i = 0
        For X = 0 To 5
            Select Case X
            Case Is = 1
                rsWithdrawal(X) = cboCustomerNo.Text
            Case Is = 5
                rsWithdrawal(X) = txtDated.Value
            Case Else
                rsWithdrawal(X) = txtDeposits(i).Text
                i = i + 1
            End Select
        Next X
        rsWithdrawal.Update
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select * FROM tblCustomers WHERE CustomerID='" & cboCustomerNo.Text & "'", cnBank, adOpenKeyset, adLockOptimistic
        If rsTemp.RecordCount > 0 Then
             rsTemp(13) = Val(lblBalance.Caption) - Val(txtDeposits(3).Text)
        Else
            Exit Sub
        End If
         rsTemp.Update
         rsTemp.Close
    
        If Check1.Value = 1 Then
            DataEnvironment1.rsCommand2.Filter = "TransactionID='" & txtDeposits(0).Text & "'"
            DataReport2.Show
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
    rsWithdrawal.AddNew
End Sub
Public Sub GenerateNewTransactCode()
    Dim lastnumber As Long, newnumber As Long
    If rsWithdrawal.BOF = True And rsWithdrawal.EOF = True Then
        lastnumber = 15000
    Else
        rsWithdrawal.MoveLast
        lastnumber = rsWithdrawal(0)
    End If
    newnumber = lastnumber + 1
    txtDeposits(0).Text = newnumber
End Sub
Private Sub cboCustomerNo_Click()
    Set rsTemp = New ADODB.Recordset
    Set rsTemp2 = New ADODB.Recordset
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
    
    rsTemp2.Open "Select * FROM tblAccTypes WHERE AccountName='" & rsTemp(12) & "'", cnBank, adOpenKeyset, adLockOptimistic
    Label3.Caption = rsTemp2(4)
    rsTemp2.Close
    rsTemp.Close
End Sub
