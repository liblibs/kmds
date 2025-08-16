Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nFocus As Integer
Dim bAutoLogin As Boolean
Dim nCount As Integer
Dim nsaveID As String

Private Sub cmdLogin_Click()
    AttemptLogin
End Sub


Private Sub AttemptLogin()
    Me.Caption = "Login " & nCount
    If gcfHandleErrors Then On Error GoTo ErrLog
    If IsNull(txtPassword) Or IsNull(txtID) Or txtID = "" Or txtPassword = "" Then
        txtID = ""
        txtPassword = ""
        txtID.SetFocus
        Beep
        Exit Sub
    End If


    If KMDSLogin(txtID, UCase(txtPassword)) Then
        nCount = 0
        txtID = ""
        txtPassword = ""
        txtID.SetFocus
        Me.Visible = False
        'DoCmd.Close acForm, Me.Name
        Call LoadDefaults
        If gcfHandleErrors Then On Error Resume Next
        PunchTheClock glnCashierID, 0
        If gcfHandleErrors Then On Error GoTo ErrLog
        LoadSwitchBoard
        Exit Sub
    Else
        nCount = nCount + 1
        txtID = ""
        txtPassword = ""
        txtID.SetFocus
    End If
ExitSub:
    Exit Sub
ErrLog:

    MsgBox "Error #" & Err.Number & " " & Err.Description, vbOKOnly, "Login:cmdLogin_Click"
    txtID = ""
    txtPassword = ""
    txtID.SetFocus
    Resume ExitSub

End Sub

Private Sub cmdLogin_GotFocus()
    AttemptLogin
End Sub

Private Sub LoadSwitchBoard()
If gcfHandleErrors Then On Error GoTo Err_Command5_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Switchboard"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    With Forms("SwitchBoard")
        .Caption = "KMDS Software © "
        .Caption = .Caption & glsCashierName & " " & gstrWindowsUserName
        .Caption = .Caption & " logged in on " & gstrComputerName & " at " & Format(gdLoginTime, "Short Time")
    End With
Exit_Command5_Click:
    Exit Sub

Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click

End Sub

Private Sub Form_Current()

    'If Not bAutoLogin Then

        'cmbItemByCode = ID
        'cmbItemByDesc = ID

    'End If
    'Me.SetFocus
    'txtID.SetFocus
End Sub

Private Sub Form_Load()
'    'force the SQL Login
'    Dim rs As Recordset
'    Set rs = CurrentDb().OpenRecordset("vwOrdhead")
'    rs.Close
'    Set rs = Nothing

End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strCashier As String
    Dim strPassword As String
    Dim strProgram As String
    Dim strParameters1 As String
    Dim strParameters2 As String
    Dim strQuit As String
    Dim strshell As String
    'Me.InsideHeight = Box6.Height + 500
    'Me.InsideWidth = Box6.Width + 500
    bAutoLogin = False
   Dim Maxit%
    If gcfHandleErrors Then
        Maxit% = ShowWindow(hWndAccessApp, 3)  'SW_SHOWMAXIMIZED)
    End If
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Dir("c:\kmds\mail\in\AutoLogin.KMDS") <> "" Then
        If gcfHandleErrors Then On Error Resume Next
        bAutoLogin = True

        Open "c:\kmds\mail\in\AutoLogin.KMDS" For Input As #1
        Line Input #1, strCashier
        'Line Input #1, strPassword ' due to kmds encryption an "EOF" is sometimes embedded in the password!
        Line Input #1, strProgram
        Line Input #1, strParameters1
        Line Input #1, strParameters2
        Line Input #1, strQuit

        Close #1

  '      If KMDSDecrypt(strPassword) <> "NoParse" Then
  '          strPassword = KMDSDecrypt(strPassword)
  '      End If

        Kill "c:\kmds\mail\in\AutoLogin.KMDS"
        If gcfHandleErrors Then On Error GoTo ErrLog

        glnCashierID = strCashier
        glsCashierName = "Autologin"


'        If KMDSLogin(CLng(strCashier), "Autologin") Then
            txtID = ""
            txtPassword = ""
            txtID.SetFocus
            Me.Visible = False
            'DoCmd.Close acForm, Me.Name
            Call LoadDefaults
            Select Case strProgram

            Case "ClubOrderCreate"
                'stDocName = "ClubOrderPrintDialog"
                'strParameters1 = "[ID]=" & Me![ClubID]
                'strParameters1 = "[ID]=" & strParameters1
                LoadSwitchBoard
                DoCmd.OpenForm strProgram
                If gcfHandleErrors Then On Error Resume Next
                Cancel = True
                'DoCmd.OpenForm "CashRegister"
                'DoCmd.Quit
            Case "ClubOrderPrintDialog"
                'stDocName = "ClubOrderPrintDialog"
                'strParameters1 = "[ID]=" & Me![ClubID]
                strParameters1 = "[ID]=" & strParameters1
                LoadSwitchBoard
                DoCmd.OpenForm strProgram, , , strParameters1, , acDialog, strParameters2
                If gcfHandleErrors Then On Error Resume Next
                Cancel = True
                '2009-09-01 the following would create errors "Not Available at this time" so i commented out
                'DoCmd.OpenForm "CashRegister"
                'Exit Sub
                'DoCmd.Quit
            Case "EndOfDay"
                If gcfHandleErrors Then On Error Resume Next

                'Call EndOFDay
                DoCmd.OpenForm "EndOFDay"

            Case "AgentTransferOut"
                'stDocName = "ClubOrderPrintDialog"
                'strParameters1 = "[ID]=" & Me![ClubID]
                'strParameters1 = "[ID]=" & strParameters1
                LoadSwitchBoard
                DoCmd.OpenForm strProgram
                If gcfHandleErrors Then On Error Resume Next
                Cancel = True
                'DoCmd.OpenForm "CashRegister"
                'DoCmd.Quit
            Case "AgentTransferIn"
                'stDocName = "ClubOrderPrintDialog"
                'strParameters1 = "[ID]=" & Me![ClubID]
                'strParameters1 = "[ID]=" & strParameters1
                LoadSwitchBoard
                DoCmd.OpenForm strProgram
                If gcfHandleErrors Then On Error Resume Next
                Cancel = True
                'DoCmd.OpenForm "CashRegister"
                'DoCmd.Quit
            Case "PurchaseOrderCreate"
                'stDocName = "ClubOrderPrintDialog"
                'strParameters1 = "[ID]=" & Me![ClubID]
                'strParameters1 = "[ID]=" & strParameters1
                LoadSwitchBoard
                DoCmd.OpenForm strProgram
                If gcfHandleErrors Then On Error Resume Next
                Cancel = True
            Case Else

                LoadSwitchBoard
            End Select
            Exit Sub
 '       End If

    Else

        If strCashier = "" Then 'no autologin
            If Get_User_Name = True Then
                txtID = gstrWindowsUserID
                txtPassword = gstrWindowsUserPassword
                AttemptLogin
            End If
        End If
    End If

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Open"
    Resume ExitSub
End Sub

Private Sub txtID_AfterUpdate()
    If txtID <> nsaveID Then
        nCount = 0
    End If
End Sub


Private Sub txtID_GotFocus()
    nFocus = 1
    nsaveID = Nz(txtID, 0)
End Sub

Private Sub txtPassword_Exit(Cancel As Integer)
'    AttemptLogin
End Sub
Private Sub cmdQuitters_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdQuitters_Click


    DoCmd.Quit

Exit_cmdQuitters_Click:
    Exit Sub

Err_cmdQuitters_Click:
    MsgBox Err.Description
    Resume Exit_cmdQuitters_Click

End Sub


Private Function HandleClick(Cap As Integer)
    Select Case Cap
    Case 10 ' clear
        txtID = ""
        txtPassword = ""
        txtID.SetFocus
        nFocus = 1
    Case 11 ' backspace
        If gcfHandleErrors Then On Error Resume Next
        Select Case nFocus
        Case 1 'id
            txtID = Left((txtID), Len(txtID) - 1)
        Case 2
            txtPassword = Left((txtPassword), Len(txtPassword) - 1)
        End Select
    Case 12 ' enter
        Select Case nFocus
        Case 1 'id
            txtPassword.SetFocus
            nFocus = 2
        Case 2
            AttemptLogin
        End Select

    Case Else 'number
        Select Case nFocus
        Case 1 'id
            txtID = Trim(txtID) & Trim(Cap)
        Case 2
            txtPassword = Trim(txtPassword) & Trim(Cap)
        End Select

    End Select

    'MsgBox Cap

End Function

Private Sub txtPassword_GotFocus()
    nFocus = 2
End Sub
