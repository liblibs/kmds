Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbsavedInfo As Boolean
Dim mbLiquorOrderGuide As Boolean



Private Sub ClubOrderGuideTemp_Enter()
    chkNonLiquor.Visible = False
    mbLiquorOrderGuide = True


End Sub

Private Sub cmbClubID_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object
If gcfHandleErrors Then On Error GoTo ErrLog

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cmbClubID], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub cmdOrderGuideTempAppend_Click()

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim stDocName As String
If gcfHandleErrors Then On Error GoTo Err_cmdOrderGuideTempAppend_Click

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderGuideTempDelete"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderGuideTempAppend"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")




    Me.ClubOrderGuideTemp.Requery
    Me.ClubOrderGuideTempRelatedSub.Requery

Exit_cmdOrderGuideTempAppend_Click:
    Me.ClubOrderGuideTemp.SetFocus
    Exit Sub

Err_cmdOrderGuideTempAppend_Click:
    MsgBox Err.Description
    Resume Exit_cmdOrderGuideTempAppend_Click

End Sub


Private Sub cmdSaveOrderGuide_Click()


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim stDocName As String
If gcfHandleErrors Then On Error GoTo ErrLog

    If gcfHandleErrors Then On Error GoTo ErrLogRB
     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderGuideDelete"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    'insert new from temp file

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderGuideAppend"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    'delete temp file
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderGuideTempDelete"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Execute
    End With
    Set comQDF = Nothing

ExitSub:
    mbsavedInfo = True
    gbUpdate = False
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrLogRB:
'    conNew.Rollback
ErrLog:
    MsgBox Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub Form_Current()
    cmbClubID = ID
    chkNonLiquor.Visible = chkNonLiquor
End Sub
Private Sub cmdInsertLine_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdInsertLine_Click

Dim stDocName As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nLineID
Dim rs As Object

    nLineID = Me.ClubOrderGuideTemp.Form!LineID

    'add 1 to all lines above
    stDocName = "qryClubOrderGuideTempIncrementLineID"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = stDocName
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters("parLineID") = nLineID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    'insert this line number
    stDocName = "qryClubOrderGuideTempInsertLineID"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = stDocName
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters("parLineID") = nLineID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


        'renumber the recordset
    Me.ClubOrderGuideTemp.Form.OrderBy = "LineID"
    Me.ClubOrderGuideTemp.Form.OrderByOn = True
    Me.ClubOrderGuideTemp.Requery

    Set rs = Me.ClubOrderGuideTemp.Form.RecordsetClone
    rs.FindFirst "[LineID] = " & Str(nLineID)
    If Not rs.EOF Then Me.ClubOrderGuideTemp.Form.Bookmark = rs.Bookmark
    rs.Close
    Set rs = Nothing




Exit_cmdInsertLine_Click:
    Me.ClubOrderGuideTemp.SetFocus
    Exit Sub

Err_cmdInsertLine_Click:
    MsgBox Err.Description
    Resume Exit_cmdInsertLine_Click

End Sub
Private Sub cmdDelete_Click()

If gcfHandleErrors Then On Error GoTo Err_cmdDelete_Click
Dim stDocName As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nLineID
Dim rs As Object

    nLineID = Me.ClubOrderGuideTemp.Form!LineID

    'add 1 to all lines above
    stDocName = "qryClubOrderGuideTempDeleteLineID"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = stDocName
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ID
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters("parLineID") = nLineID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    Me.ClubOrderGuideTemp.Requery

    'renumber the recordset
    Me.ClubOrderGuideTemp.Form.OrderBy = "LineID"
    Me.ClubOrderGuideTemp.Form.OrderByOn = True

    'Me.ClubOrderGuideTempRelatedSub.Form.OrderBy = "LineID"
    'Me.ClubOrderGuideTempRelatedSub.Form.OrderByOn = True

    Set rs = Me.ClubOrderGuideTemp.Form.RecordsetClone
    With rs
        Do While Not .EOF
            .Edit
            !LineID = .AbsolutePosition + 1
            .Update
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing

    Me.ClubOrderGuideTemp.Requery
    Set rs = Me.ClubOrderGuideTemp.Form.RecordsetClone
    rs.FindFirst "[LineID] = " & Str(nLineID)
    If Not rs.EOF Then Me.ClubOrderGuideTemp.Form.Bookmark = rs.Bookmark
    rs.Close
    Set rs = Nothing

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
End Sub
Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click

    DoCmd.Close


Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not gcfHandleErrors Then Debug.Print KeyCode
    Select Case KeyCode
    Case 45 'insert

    Case 46 'delete
    Case Else
    End Select
End Sub

Private Sub Form_Open(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        Me.InsideHeight = Me.FormHeader.Height + Box0.Height + Me.FormFooter.Height + 500
        Me.InsideWidth = Box0.Width + 250
        mbsavedInfo = False
        gbUpdate = False
        cmdOrderGuideTempAppend_Click
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Open"
    Resume ExitSub
End Sub

Private Sub Form_Unload(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click
    Dim nResp As Integer
    If gbUpdate = True Then
        nResp = MsgBox("By closing this for you will lose any changes you have made" & vbCrLf & "Do you want to close this form?", vbYesNo)
        Select Case nResp
        Case vbNo
            Cancel = True

        Case Else
        End Select
    End If
Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
End Sub
