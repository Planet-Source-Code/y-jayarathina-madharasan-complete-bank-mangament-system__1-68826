Attribute VB_Name = "Module1"
'Module1
Public cnBank As Connection
Public rsCustomers As Recordset
Public rsDeposit As Recordset
Public rsWithdrawal As Recordset
Public rsAccTypes As Recordset

Public rsTemp As Recordset
Public rsTemp2 As Recordset

Public X As Integer

Public Sub create_navigation_buttons(frm As Form)
    With frm
        With .cmdOptions(0)
            .Picture = LoadPicture(App.Path & "\pictures\n1.jpg")
            .DisabledPicture = LoadPicture(App.Path & "\pictures\nd1.jpg")
        End With
        For i = 1 To 4
            Load .cmdOptions(i)
            With .cmdOptions(i)
                .Visible = True
                .Left = frm.cmdOptions(i - 1).Left + frm.cmdOptions(i - 1).Width + 25
                .Picture = LoadPicture(App.Path & "\pictures\n" & (i + 1) & ".jpg")
                .DisabledPicture = LoadPicture(App.Path & "\pictures\nd" & (i + 1) & ".jpg")
            End With
        Next i
        
        .cmdNavigate(0).Picture = LoadPicture(App.Path & "\pictures\n6.jpg")
        For i = 1 To 3
            Load .cmdNavigate(i)
            With .cmdNavigate(i)
                .Visible = True
                .Left = frm.cmdNavigate(i - 1).Left + frm.cmdNavigate(i - 1).Width + 15
                .Picture = LoadPicture(App.Path & "\pictures\n" & (i + 6) & ".jpg")
            End With
        Next i
        .cmdOptions(1).Enabled = False: .cmdOptions(2).Enabled = False
    End With
End Sub

Public Sub create_countries(combobx As ComboBox)
    Dim st() As String
    Dim st2 As String
    Dim i As Integer
    st2 = "United Kingdom\Albania\Argentina\Afghanistan\Algeria\Australia\Belgium\Brazil\China\Canada\Colombia\Costa Rica\Czech Republic\Germany\Denmark\Egypt\Ecuador\United Arab Emirates\Finland\France\Greece\Hong Kong\Hungary\Indonesia\Ireland\India\Israel\Italy\Japan\Lebanon\Malaysia\Mexico\Netherlands\Norway\New Zealand\Austria\Philippines\Pakistan\Poland\Portugal\Peru\Puerto Rico\Russia\Saudi Arabia\Sweden\Spain\Singapore\Switzerland\Thailand\Turkey\Taiwan\Tajikistan\Tanzania\Tunisia\Tuvalu\United States of America\Ukraine\Venezuela\South Africa\Uganda\Uzbekistan\Vatican City\Vietnam\Yemen\Zimbabwe"
    st = Split(st2, "\")
    For i = LBound(st) To UBound(st)
        combobx.AddItem (st(i))
    Next i
    combobx.ListIndex = 25
End Sub

Public Sub connectDatabase()
    Set cnBank = New ADODB.Connection
    
    With cnBank
        .Provider = "Microsoft.JET.OLEDB.4.0"
        .ConnectionString = App.Path & "\dbBank.mdb"
        .Open
    End With
    
    Set rsCustomers = New ADODB.Recordset
    rsCustomers.Open "tblCustomers", cnBank, adOpenKeyset, adLockOptimistic
    
    Set rsDeposit = New ADODB.Recordset
    rsDeposit.Open "tblDeposits", cnBank, adOpenKeyset, adLockOptimistic
    
    Set rsWithdrawal = New ADODB.Recordset
    rsWithdrawal.Open "tblWithdrawals", cnBank, adOpenKeyset, adLockOptimistic
    
    Set rsAccTypes = New ADODB.Recordset
    rsAccTypes.Open "tblAccTypes", cnBank, adOpenKeyset, adLockOptimistic

End Sub

Public Sub clear_Form_Controls(frm As Form)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Text = ""
        End If
    Next ctrl
End Sub
Public Sub Messager()
    MsgBox "Please Ensure that all fields are Complete", vbExclamation
End Sub
Public Sub Lock_Form_Controls(frm As Form)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        ElseIf TypeOf ctrl Is DTPicker Then
            ctrl.Enabled = False
        End If
    Next ctrl
End Sub
Public Sub UnLock_Form_Controls(frm As Form)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = False
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Locked = False
        ElseIf TypeOf ctrl Is DTPicker Then
            ctrl.Enabled = True
        End If
    Next ctrl
End Sub
Public Sub MoveToFirst(rsFirst As Recordset)
    With rsFirst
        Call CheckDatabaseStatus(rsFirst)
        .MoveFirst
        If .BOF Then
            .MoveFirst
            MsgBox "This is the first Record..", vbInformation
            Exit Sub
        End If
    End With
End Sub
Public Sub MoveToPrev(rsPrev As Recordset)
    With rsPrev
        Call CheckDatabaseStatus(rsPrev)
        .MovePrevious
        If .BOF Then
            .MoveFirst
            MsgBox "This is the first Record..", vbInformation
            Exit Sub
        End If
    End With
End Sub
Public Sub MoveToNext(rsNext As Recordset)
    With rsNext
        Call CheckDatabaseStatus(rsNext)
        .MoveNext
        If .EOF Then
            .MoveLast
            MsgBox "This is the last Record..", vbInformation
            Exit Sub
        End If
    End With
End Sub

Public Sub MoveToLast(rsLast As Recordset)
    With rsLast
        Call CheckDatabaseStatus(rsLast)
        .MoveLast
        If .EOF Then
            .MoveLast
            MsgBox "This is the last Record..", vbInformation
            Exit Sub
        End If
    End With
End Sub
Public Sub CheckDatabaseStatus(rsStat As Recordset)
    With rsStat
        If .BOF = True And .EOF = True Then
            MsgBox "There are currently No records Available for this module", vbInformation
            Exit Sub
        End If
    End With
End Sub
