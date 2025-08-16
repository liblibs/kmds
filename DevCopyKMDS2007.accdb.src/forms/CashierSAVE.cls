Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean

Private Sub cboID_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cboID], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdfrmAddNew

    DoCmd.GoToRecord , , acNewRec

Exit_cmdfrmAddNew:
    Exit Sub

Err_cmdfrmAddNew:
    MsgBox Err.Description
    Resume Exit_cmdfrmAddNew
End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    'set focus to combo
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

    'refresh combo box,
'   cboID.Requery
'   cboID = ID
'   cboID.Visible = True
'   cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmSaveRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmUndo_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub
End Sub

Private Sub cmdSetAllowed_Click()

Dim nResp As Integer
Dim stDocName As String
Dim stLinkCriteria As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    If gcfHandleErrors Then On Error GoTo ErrLog

    nResp = MsgBox("Set ALL Allowed flags to...", vbYesNoCancel)
    Select Case nResp
    Case vbYes
        nResp = -1
    Case vbNo
        nResp = 0
    Case Else
        Exit Sub
    End Select

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCashierAndProgramsSetAllowedFlag"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parCashierID", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("parAllowed", adBoolean, adParamInput)
        .Parameters("parCashierID") = ID
        .Parameters("parAllowed") = nResp
        .Execute
    End With
    Me.CashiersAndProgramsDatasheet.Requery
    With Me.CashiersAndPrograms
        .Requery
        .Form.OrderBy = "Lookup_cmbProgramName.IT"
        .Form.OrderByOn = True
        'Set rs = .Form.Recordset.Clone
        'rs.FindFirst "[ProgramName] = '" & cmdSetAllowedForThisProgram.Tag & "'"
        'If Not rs.EOF Then .Form.Bookmark = rs.Bookmark
    End With
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdSetAllowed_Click"
    Resume ExitSub
End Sub



Private Sub cmdSetAllowedForThisProgram_Click()
Dim nResp As Integer
Dim stDocName As String
Dim stLinkCriteria As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")
Dim rs
    If gcfHandleErrors Then On Error GoTo ErrLog

    nResp = MsgBox("Set ALL Allowed flags to...", vbYesNoCancel)
    Select Case nResp
    Case vbYes
        nResp = -1
    Case vbNo
        nResp = 0
    Case Else
        Exit Sub
    End Select

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCashierAndProgramsSetAllowedFlagByProgram"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parAllowed", adBoolean, adParamInput)
        .Parameters.Append .CreateParameter("parProgramName", adBSTR, adParamInput)
        .Parameters("parProgramName") = cmdSetAllowedForThisProgram.Tag
        .Parameters("parAllowed") = nResp
        .Execute
    End With
    With Me.CashiersAndPrograms
        .Requery
        .Form.OrderBy = "Lookup_cmbProgramName.IT"
        .Form.OrderByOn = True
        'Set rs = .Form.Recordset.Clone
        'rs.FindFirst "[ProgramName] = '" & cmdSetAllowedForThisProgram.Tag & "'"
        'If Not rs.EOF Then .Form.Bookmark = rs.Bookmark
    End With


    '
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdSetAllowed_Click" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub Form_AfterInsert()
    Dim stDocName As String
    Dim nsaveID As Integer

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    nsaveID = ID
    stDocName = "qryCashiersAndProgramsAppend"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    Me.Requery
    cboID.Requery

        ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(nsaveID, 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark


    'form current fires so you don't have to do the next command
    'cboID = ID
    SetOptions True
    'CurrentDb().Execute "qryProgramAppend"
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdSetAllowed_Click" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub



Private Sub Form_BeforeUpdate(Cancel As Integer)
    gbRunMailChanges = True
    ChangeDate = Now
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.NewRecord Then
        ID.Visible = True
        ID.SetFocus
        cboID.Visible = False
    Else
        cboID.Visible = True
        ID.Visible = False
        cboID = ID
        cmdSetAllowed.Caption = "Set &ALL Allowed Flags for #" & ID & " " & cboID.Column(1)
    End If

    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = False
    cmdfrmUndo.Visible = False

    'Me.CashiersAndPrograms.Form.OrderBy = "Lookup_cmbProgramName.It"
    'Me.CashiersAndPrograms.Form.OrderByOn = True




ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_Load()
    gbRunMailChanges = False

End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim stDocName As String


'    SetOptions False
'    stDocName = "qryProgramAppend"
'    DoCmd.OpenQuery stDocName, acNormal, acEdit
'    stDocName = "qryCashiersAndProgramsAppend"
'
'
'    DoCmd.OpenQuery stDocName, acNormal, acEdit
'    Me.Requery
'    SetOptions True

    'CurrentDb().Execute "qryProgramAppend"
   ' Cancel = FormLogin(glnCashierID, Me)


End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then
        MailChanges (False)
    End If
End Sub
Private Sub cmdAppendCashiersAndPrograms_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAppendCashiersAndPrograms_Click

    Dim stDocName As String
    SetOptions False
    stDocName = "qryProgramAppend"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qryCashiersAndProgramsAppend"


    DoCmd.OpenQuery stDocName, acNormal, acEdit
    Me.Requery

    SetOptions True

Exit_cmdAppendCashiersAndPrograms_Click:
    Exit Sub

Err_cmdAppendCashiersAndPrograms_Click:
    MsgBox Err.Description
    Resume Exit_cmdAppendCashiersAndPrograms_Click

End Sub
Private Sub cmdCashRegisterFlags_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdCashRegisterFlags_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "CashierRegisterFlags"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCashRegisterFlags_Click:
    Exit Sub

Err_cmdCashRegisterFlags_Click:
    MsgBox Err.Description
    Resume Exit_cmdCashRegisterFlags_Click

End Sub
Private Sub cmdCashiersBackOfficeFlags_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdCashiersBackOfficeFlags_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "CashiersBackOfficeFlags"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdCashiersBackOfficeFlags_Click:
    Exit Sub

Err_cmdCashiersBackOfficeFlags_Click:
    MsgBox Err.Description
    Resume Exit_cmdCashiersBackOfficeFlags_Click

End Sub

Private Sub cmdCRViewer_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdCRViewer_Click

    Dim stAppName As String

    stAppName = "\KMDS\CRView.exe"
    Call Shell(stAppName, 1)

Exit_cmdCRViewer_Click:
    Exit Sub

Err_cmdCRViewer_Click:
    MsgBox Err.Description
    Resume Exit_cmdCRViewer_Click

End Sub
