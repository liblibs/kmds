Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_Open(Cancel As Integer)

    '10/3/2004 added bOKToRunRegister to make sure machine is configured to cash register

    Dim bOKToRunRegister As Boolean
    If gcfHandleErrors Then On Error Resume Next
    txtCashRegisterNumber = glnCashsysID
    txtCashRegisterName = gstrComputerName

    bOKToRunRegister = True

    Select Case glnCashsysID
    Case 1
        If Cr1Installed = 0 Or Cr1Online = 0 Then bOKToRunRegister = False
    Case 2
        If Cr2Installed = 0 Or Cr2Online = 0 Then bOKToRunRegister = False
    Case 3
        If Cr3Installed = 0 Or Cr3Online = 0 Then bOKToRunRegister = False
    Case 4
        If Cr4Installed = 0 Or Cr4Online = 0 Then bOKToRunRegister = False
    Case 5
        If Cr5Installed = 0 Or Cr5Online = 0 Then bOKToRunRegister = False
    Case Else
        bOKToRunRegister = False
    End Select

    Cancel = True
    If Not bOKToRunRegister Then
        MsgBox "This machine (register #" & glnCashsysID & " " & txtCashRegisterName & ") is not allowed to run cash register!"
    Else
        Open "c:\kmds\mail\in\AutoLogin.cr" For Output As #1
        Print #1, glnCashierID
        Close #1
        Shell "c:\kmds\cr3.exe"
        DoCmd.Quit
    End If
End Sub
