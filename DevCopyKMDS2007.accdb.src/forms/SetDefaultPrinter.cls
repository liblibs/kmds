Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'note Vista installs the prnmngr scripts in C:\Windows\System32\Printing_Admin_Scripts\en-US
'            the *.scr files have to be copied to C:\Windows\System32\

'Microsoft (R) Windows Script Host Version 5.7
'Copyright (C) Microsoft Corporation. All rights reserved.
'
'Usage: prnmngr [-adxgtl?][c] [-s server][-p printer][-m driver model]
'               [-r port][-u user name][-w password]
'Arguments:
'-a     - add local printer
'-ac    - add printer connection
'-d     - delete printer
'-g     - get the default printer
'-l     - list printers
'-m     - driver model
'-p     - printer name
'-r     - port name
'-s     - server name
'-t     - set the default printer
'-u     - user name
'-w     - password
'-x     - delete all printers
'-xc    - delete all printer connections
'-xo    - delete all local printers
'-?     - display command usage
'
'Examples:
'prnmngr -a -p "printer" -m "driver" -r "lpt1:"
'prnmngr -d -p "printer" -s server
'prnmngr -ac -p "\\server\printer"
'prnmngr -d -p "\\server\printer"
'prnmngr -x -s server
'prnmngr -xo
'prnmngr -l -s server
'prnmngr -g
'prnmngr -t -p "\\server\printer"

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    SetKMDSSettings "Printer_PortName", txtPortName
    SetKMDSSettings "Printer_DriverModel", txtDriverModel
    SetKMDSSettings "Printer_PrinterName", txtPrinterName
    If InStr(txtPrinterName, txtComputerName) Then
        lblLocalPrinter.Visible = True
    Else
        lblLocalPrinter.Visible = False
    End If

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdfrmSaveRecord_Click"
    Resume ExitSub

End Sub

Private Sub cmdLockSettings_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    txtPortName.Locked = True
    txtDriverModel.Locked = True
    txtPrinterName.Locked = True
    txtPrinterName.SetFocus
    cmdUnlockSettings.Visible = True
    cmdLockSettings.Visible = False
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdLockSettings_Click"
    Resume ExitSub
End Sub

Private Sub cmdSetPrinter_Click()
    Dim fsoTemp As New FileSystemObject
    Dim txtstr As TextStream
If gcfHandleErrors Then On Error GoTo ErrLog
    If fsoTemp.FileExists("c:\kmds\SetDefaultPrinter.bat") Then fsoTemp.DeleteFile "c:\kmds\SetDefaultPrinter.Bat"
    Set txtstr = fsoTemp.OpenTextFile("c:\kmds\SetDefaultPrinter.Bat", ForWriting, True)
    With txtstr
        'cscript //h:cscript //s
        'prnmngr -ac -p "KMDSReports" -m "Brother HL-2140 series" -p "\\BO1\BrotherH"
        'prnmngr -t -p "\\BO1\KMDS"
        .WriteLine "cscript //h:cscript //s"
        If lblLocalPrinter.Visible Then ' local printer just set it to default
            .WriteLine "prnmngr -t -p " & Chr(34) & Replace(txtPrinterName, "\\" & txtComputerName & "\", "") & Chr(34) 'remove UNC from name
        Else 'Network printer
            .WriteLine "prnmngr -ac -m " & Chr(34) & txtDriverModel & Chr(34) & " -p " & Chr(34) & txtPrinterName & Chr(34)
            .WriteLine "prnmngr -t -p " & Chr(34) & txtPrinterName & Chr(34)
        End If
    End With
    txtstr.Close
    Set fsoTemp = Nothing
    Shell "c:\kmds\SetDefaultPrinter.Bat", vbNormalFocus

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdSetPrinter_Click"
    Resume ExitSub
End Sub

Private Sub cmdShowPrinters_Click()
    Dim fsoTemp As New FileSystemObject
    Dim txtstr As TextStream
If gcfHandleErrors Then On Error GoTo ErrLog
    If fsoTemp.FileExists("c:\kmds\ShowPrinters.Txt") Then fsoTemp.DeleteFile "c:\kmds\ShowPrinters.Txt"
    If fsoTemp.FileExists("c:\kmds\ShowPrinters.bat") Then fsoTemp.DeleteFile "c:\kmds\ShowPrinters.Bat"
    Set txtstr = fsoTemp.OpenTextFile("c:\kmds\ShowPrinters.Bat", ForWriting, True)
    With txtstr
        .WriteLine "cscript //h:cscript //s"
        .WriteLine "prnmngr -l > c:\kmds\ShowPrinters.Txt"
        .WriteLine "start notepad c:\kmds\ShowPrinters.Txt"
    End With
    txtstr.Close
    Set fsoTemp = Nothing
    Shell "c:\kmds\ShowPrinters.Bat", vbNormalFocus

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdShowPrinters_Click"
    Resume ExitSub
End Sub

Private Sub cmdUnlockSettings_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    txtPortName.Locked = False
    txtDriverModel.Locked = False
    txtPrinterName.Locked = False
    txtPrinterName.SetFocus
    cmdUnlockSettings.Visible = False
    cmdLockSettings.Visible = True

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "cmdUnlockSettings_Click"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    Cancel = FormLogin(glnCashierID, Me)
    txtComputerName = gstrComputerName
    txtUser = glsCashierName
    txtPortName = GetKMDSSettings("Printer_PortName", "")
    txtDriverModel = GetKMDSSettings("Printer_DriverModel", "")
    txtPrinterName = GetKMDSSettings("Printer_PrinterName", "")
    If InStr(txtPrinterName, txtComputerName) Then
        lblLocalPrinter.Visible = True
    Else
        lblLocalPrinter.Visible = False
    End If


ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "Form_Open"
    Resume ExitSub
End Sub


Private Sub txtDriverModel_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "txtDriverModel_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub txtPortName_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "txtPortName_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub txtPrinterName_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "txtPrinterName_AfterUpdate"
    Resume ExitSub
End Sub
