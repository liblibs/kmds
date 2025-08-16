Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strFnameInput As String
Dim strfnameBatch As String
Dim fsoTemp As New FileSystemObject
Dim strWorkStationList As String
Dim txtstr As TextStream

Private Sub Command8_Click()
    txtPwd = UCase(txtPwd)
    txtPwd1 = UCase(txtPwd1)
    If StrComp(txtPwd, txtPwd1, 0) <> 0 Then
        MsgBox "Passwords to not match"
        txtPwd.SetFocus

        Exit Sub
    End If

    If KMDSPasswordCheck(txtPwd) Then

        'SQLPassThru ("sp_CashierPasswordHistory " & CStr(glnCashierIDForPasswordForm) & ", " & KMDSEncrypt(txtPwd))
        MsgBox "Password has been reset!"
        gbRunMailChanges = True

        DoCmd.Close acForm, Me.Name
    Else
        txtPwd.SetFocus
    End If

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Command8_Click"
    Resume ExitSub

End Sub


Private Sub Form_Load()
    Me.InsideHeight = boxScreen.Height + 300
    Me.InsideWidth = boxScreen.Width + 300

End Sub

Private Sub Form_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog

    If gstrCashsysInsLoc = "" Then LoadDefaults

'    strWorkStationList = CashStatus()

'    If InStr(strWorkStationList, "Problems") Then
'        MsgBox "Not all active workstations are connected to network" & vbcrlf & "Press enter to see which one"
'        MsgBox strWorkStationList
'        MsgBox "Connect all active workstations to network and try again."
'        Cancel = True
'    End If
    txtUser = glnCashierIDForPasswordForm

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Form_Open"
    Resume ExitSub

End Sub
