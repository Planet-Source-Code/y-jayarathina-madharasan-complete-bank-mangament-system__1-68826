VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTransactions 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transactions"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10485
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dated"
      Height          =   1095
      Left            =   5400
      TabIndex        =   11
      Top             =   1200
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "Proceed"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   71041025
         CurrentDate     =   38311
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   71041025
         CurrentDate     =   38311
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
      Begin VB.Label dtToj 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Choose the View Mode"
      Height          =   615
      Left            =   3675
      TabIndex        =   9
      Top             =   480
      Width           =   2655
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "View All"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Custom"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "By Coustomer Details"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
      Begin VB.ComboBox cboCustomerID 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Text            =   "Select..."
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cboFirst 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "Select..."
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cboAccNo 
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Text            =   "Select..."
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2295
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Deposit Details:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Withdrawal Details:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -120
      TabIndex        =   19
      Top             =   5280
      Width           =   10575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -120
      TabIndex        =   8
      Top             =   2280
      Width           =   10575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Caption         =   "Transaction Statement"
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
      TabIndex        =   7
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmTransactions
Dim lstItem As ListItem
Private Sub Command1_Click()
    If find_tran("tblWithdrawals Where Dated BETWEEN #" & dtFrom.Value & "# AND #" & dtTo.Value & "#", MSFlexGrid1, 1) = False And find_tran("tblDeposits Where Dated BETWEEN #" & dtFrom.Value & "# AND #" & dtTo.Value & "#", MSFlexGrid2, 2) = False Then
        MsgBox "No Transaction found. Please Try Again", vbInformation
    End If
End Sub
Private Sub set_dat(qr As String)
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select * from " & qr, cnBank, adOpenKeyset, adLockOptimistic
    With rsTemp
        If .RecordCount > 0 Then
            cboAccNo = !AccountNo
            cboCustomerID = !CustomerID
            cboFirst = !FirstName
        Else
            MsgBox "Invalid customer ID/Name/Account NO. Please Try Again", vbInformation
            Exit Sub
        End If
        .Close
    End With
    If find_tran("tblWithdrawals Where AccountNo='" & cboAccNo.Text & "'", MSFlexGrid1, 1) = False And find_tran("tblDeposits Where AccountNo='" & cboAccNo.Text & "'", MSFlexGrid2, 2) = False Then
        MsgBox "No Transaction found. Please Try Again", vbInformation
    End If
End Sub


Private Sub cboCustomerID_Click()
    Call set_dat("tblCustomers Where CustomerID='" & cboCustomerID.Text & "'")
End Sub

Private Sub cboCustomerID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cboCustomerID_Click
End Sub

Private Sub cboFirst_Click()
    Call set_dat("tblCustomers Where FirstName='" & cboFirst.Text & "'")
End Sub
Private Sub cboAccNo_Click()
    Call set_dat("tblCustomers Where AccountNo='" & cboAccNo.Text & "'")
End Sub

Function find_tran(tbl As String, MSFlex As MSFlexGrid, pu As Integer) As Boolean
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select * from " & tbl, cnBank, adOpenKeyset, adLockOptimistic
    MSFlex.Clear
    If pu = 1 Then
        Call create_with(rsTemp)
    Else
        Call create_dep(rsTemp)
    End If
    With rsTemp
        If .RecordCount > 0 Then
            Call LoadListView(rsTemp, MSFlex)
        Else
            find_tran = False
            Exit Function
        End If
        .Close
    End With
    find_tran = True
End Function

Private Sub cboAccNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cboAccNo_Click
End Sub

Private Sub cboFirst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cboFirst_Click
End Sub

Private Sub Form_Load()
    Option1(1).Value = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    Call connectDatabase
    Call create_with(rsWithdrawal)
    Call LoadListView(rsWithdrawal, MSFlexGrid1)
    Call create_dep(rsDeposit)
    Call LoadListView(rsDeposit, MSFlexGrid2)
    With rsCustomers
        .MoveFirst
        For X = 1 To .RecordCount
            cboCustomerID.AddItem !CustomerID
            cboFirst.AddItem !FirstName
            cboAccNo.AddItem !AccountNo
            .MoveNext
        Next X
    End With
    Frame1.Enabled = False
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Frame1.Enabled = True
        Frame2.Enabled = True
    Else
        Frame1.Enabled = False
        Frame2.Enabled = False
        If find_tran("tblWithdrawals", MSFlexGrid1, 1) = False And find_tran("tblDeposits", MSFlexGrid2, 2) = False Then
            MsgBox "No Transaction found.", vbInformation
        End If
        cboAccNo.Text = "Select..."
        cboCustomerID.Text = "Select..."
        cboFirst.Text = "Select..."
    End If
End Sub
Public Sub create_dep(rs As Recordset)
    With MSFlexGrid2
        .Rows = rs.RecordCount + 1
        .Cols = 8
        .Row = 0:
        .Col = 0: .Text = "TransactionID"
        .Col = 1: .Text = "CustomerID"
        .Col = 2: .Text = "AccountNo"
        .Col = 3: .Text = "Narration"
        .Col = 4: .Text = "AmountDeposited"
        .Col = 5: .Text = "Mode"
        .Col = 6: .Text = "CheckNo"
        .Col = 7: .Text = "Dated"
    End With
End Sub
Public Sub create_with(rs As Recordset)
    With MSFlexGrid1
        .Rows = rs.RecordCount + 1
        .Cols = 6
        .Row = 0
        .Col = 0: .Text = "TransactionID"
        .Col = 1: .Text = "CustomerID"
        .Col = 2: .Text = "AccountNo"
        .Col = 3: .Text = "Narration"
        .Col = 4: .Text = "AmountWithdrawn"
        .Col = 5: .Text = "Dated"
    End With

End Sub
Public Sub LoadListView(myrs As Recordset, MSFlex As MSFlexGrid)
    MSFlex.Rows = myrs.RecordCount + 1
    MSFlex.Row = 0
    With myrs
        While Not .EOF
            MSFlex.Row = MSFlex.Row + 1
                For i = 0 To .Fields.Count - 1
                    MSFlex.Col = i
                    MSFlex.Text = .Fields.Item(i)
                Next i
            .MoveNext
        Wend
    End With
End Sub
