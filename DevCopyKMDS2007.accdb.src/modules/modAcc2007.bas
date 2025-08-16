Attribute VB_Name = "modAcc2007"
Option Compare Database

Public Function DoMouseWheel(frm As Form, lngCount As Long) As Integer
On Error GoTo Err_Handler
    'Purpose:   Make the MouseWheel scroll in Form View in Access 2007.
    '           This code lets Access 2007 behave like older versions.
    'Return:    1 if moved forward a record, -1 if moved back a record, 0 if not moved.
    'Author:    Allen Browne, February 2007.
    'Usage:     In the MouseWheel event procedure of the form:
    '               Call DoMouseWheel(Me, Count)
    Dim strMsg As String
    'Run this only in Access 2007 and later, and only in Form view.
    If (Val(SysCmd(acSysCmdAccessVer)) >= 12#) And (frm.CurrentView = 1) And (lngCount <> 0&) Then
        'Save any edits before moving record.
        RunCommand acCmdSaveRecord
        'Move back a record if Count is negative, otherwise forward.
        RunCommand IIf(lngCount < 0&, acCmdRecordsGoToPrevious, acCmdRecordsGoToNext)
        DoMouseWheel = Sgn(lngCount)
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 2046&                 'Can't move before first, after last, etc.
        Beep
    Case 3314&, 2101&, 2115&   'Can't save the current record.
        strMsg = "Cannot scroll to another record, as this one can't be saved."
        MsgBox strMsg, vbInformation, "Cannot scroll"
    Case Else
        strMsg = "Error " & Err.Number & ": " & Err.Description
        MsgBox strMsg, vbInformation, "Cannot scroll"
    End Select
    Resume Exit_Handler
End Function


Public Function DatasheetView(frm As Form)
    DoCmd.OpenForm frm.Name, acFormDS
End Function


Public Function PrintPreView(frm As Form)
    'set the form's tag to the name of the listing report..
    If gcfHandleErrors Then On Error GoTo ErrLog
    DoCmd.OpenReport frm.Tag, acViewPreview, , , acWindowNormal
ExitFunc:
    Exit Function
ErrLog:
    MsgBox "   set the form's tag to the name of the listing report.."
    Resume ExitFunc
End Function


Public Function ExportReport(frm As String)
    DoCmd.OpenReport frm, acViewPreview
    'DoCmd.OutputTo acOutputQuery, Reports(0).RecordSource, acFormatXLS, "c:\test.xls", True
    DoCmd.Close acReport, frm


End Function


Public Function GetKMDSSettings(KeyName As String, DefaultValue As String) As String
    'KMDS2007 uses table "Settings" to store matched pairs in a Key/Value schema..
    'if the keyname cannot be found then a new key is setup using the default value..
    'the table can be accessed through the g-5 screen under 'settings' tab.

    Dim strSQL As String


'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")



    If gcfHandleErrors Then On Error GoTo ErrLog
    With rstemp
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "Settings", CurrentProject.Connection, , , adCmdTable
    End With

    rstemp.Find "KeyName='" & KeyName & "'"  ', adSeekFirstEQ

    With rstemp
        If .EOF Then
            .AddNew
            !KeyName = KeyName
            !Value = DefaultValue
            .Update
            GetKMDSSettings = DefaultValue
            MsgBox "New Key " & vbCrLf & KeyName & " : " & DefaultValue
            gbRunMailChanges = True
        Else
            If Not IsNull(!Value) Then GetKMDSSettings = !Value
            'If Not gcfHandleErrors Then Debug.Print KeyName; " = "; Trim(!Value)
        End If
        GetKMDSSettings = Trim(GetKMDSSettings)
    End With
    rstemp.Close
    Set rstemp = Nothing
ExitFunc:
    Exit Function
ErrLog:
    MsgBox "error #" & Err.Number & " " & Err.Description, , "GetKMDSEmailSettings"
    Resume ExitFunc
End Function


Public Function SetKMDSSettings(KeyName As String, Value As String) As String
    'KMDS2007 uses table "Settings" to store matched pairs in a Key/Value schema..
    'if the keyname cannot be found then a new key is setup using the default value..
    'the table can be accessed through the g-5 screen under 'settings' tab.
    Dim strSQL As String

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    If gcfHandleErrors Then On Error GoTo ErrLog
    With rstemp
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "Settings", CurrentProject.Connection, , , adCmdTable
    End With

    rstemp.Find "KeyName='" & KeyName & "'"  ', adSeekFirstEQ

    With rstemp
        If .EOF Then
            .AddNew
            !KeyName = KeyName
            MsgBox "New Key " & vbCrLf & KeyName & " : " & Value

        End If
        !Value = Value
        .Update
        SetKMDSSettings = Value
    End With
    rstemp.Close
    Set rstemp = Nothing
ExitFunc:
    Exit Function
ErrLog:
    MsgBox "error #" & Err.Number & " " & Err.Description, , "SetKMDSSettings"
    Resume ExitFunc
End Function




Public Sub Add12MonthsToClubHistoryTable()
Dim nI As Integer
Dim strSQL As String
    For nI = 1 To 12
        strSQL = "INSERT INTO ClubHistoryByBottleByMonth ( [Year-Month] ) "
        strSQL = strSQL & "SELECT '" & Format(nI, "00") & "' AS [Year-Month]"
        DoCmd.RunSQL strSQL, False
    Next nI

End Sub
