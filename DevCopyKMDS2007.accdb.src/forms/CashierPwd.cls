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
    Dim strB As String
    Dim strAdUser As String
    Dim strAdPwd As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    If txtPwd = "" Then
        Beep
        MsgBox "Password can not be blank"
        Exit Sub
    End If

    If txtPwd <> txtPwd1 Then
        MsgBox "Passwords to not match"
        Exit Sub
    End If
    strAdUser = GetKMDSSettings("AdUser", "skelly")
    strAdPwd = GetKMDSSettings("AdPwd", "Kmds11605")

    Set txtstr = fsoTemp.OpenTextFile("C:\KMDS\ws.inp", ForWriting, True, TristateUseDefault)
    txtstr.Write strWorkStationList
    txtstr.Close
    Set txtstr = Nothing

    Set txtstr = fsoTemp.OpenTextFile("c:\KMDS\psadduser.bat", ForWriting, True, TristateUseDefault)
    txtstr.WriteLine "@echo off"
    txtstr.WriteLine "c:\KMDS\psexec   @%5 -u %3 -p %4 net user %1 /ADD > c:\KMDS\psadduser.doc"
    txtstr.WriteLine "c:\KMDS\pspasswd @%5 -u %3 -p %4 %1 %2 >> c:\KMDS\psadduser.doc"
    txtstr.WriteLine "c:\KMDS\psexec   @%5 -u %3 -p %4 net localgroup " & Chr(34) & "Power Users" & Chr(34) & " %1 /ADD >> c:\KMDS\psadduser.doc"
    If optAdministrator Then
        txtstr.WriteLine "c:\KMDS\psexec   @%5 -u %3 -p %4 net localgroup " & Chr(34) & "Administrators" & Chr(34) & " %1 /ADD >> c:\KMDS\psadduser.doc"
    Else
        txtstr.WriteLine "c:\KMDS\psexec   @%5 -u %3 -p %4 net localgroup " & Chr(34) & "Administrators" & Chr(34) & " %1 /DELETE >> c:\KMDS\psadduser.doc"
    End If
    txtstr.Close
    Set txtstr = Nothing

    strB = "c:\KMDS\psadduser.bat " & txtUser & " " & txtPwd & " " & strAdUser & " " & strAdPwd & " C:\KMDS\ws.inp"

    'MsgBox strB
    'Shell strB, vbHide
    WaitForExecCmd strB
    DoCmd.Close
    'fsotemp.DeleteFile "c:\KMDS\psadduser.bat", True

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

    strWorkStationList = CashStatus()

    If InStr(strWorkStationList, "Problems") Then
        MsgBox "Not all active workstations are connected to network" & vbCrLf & "Press enter to see which one"
        MsgBox strWorkStationList
        MsgBox "Connect all active workstations to network and try again."
        Cancel = True
    End If
    txtUser = glnCashierIDForPasswordForm

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Form_Open"
    Resume ExitSub

End Sub
