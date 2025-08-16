Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboAgentByID_AfterUpdate()
   ' Find the record that matches the control.
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[AgentNumber] = " & Str(Me![cboAgentByID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cboID_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cboID_AfterUpdate()
   ' Find the record that matches the control.
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[AgentNumber] = " & Str(Me![cboID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cboID_AfterUpdate"
    Resume ExitSub

End Sub



Private Sub Check49_AfterUpdate()

End Sub

Private Sub cmdActivate_Click()
    DoCmd.OpenForm "AgentsListBox"
    DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdfrmAddNew

    DoCmd.GoToRecord , , acNewRec

Exit_cmdfrmAddNew:
    Exit Sub

Err_cmdfrmAddNew:
    EMsgBox Me, "cmdfrmAddNew_Click"
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

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70
    Me.Requery


ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501
        MsgBox "There are related items for this record"
    Case Else
        EMsgBox Me, "cmdfrmDeleteRecord_Click"
    End Select
    Resume ExitSub
End Sub

Private Sub cmdfrmSaveRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    cmdfrmAddNew.Visible = True
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

Private Sub Combo53_AfterUpdate()

End Sub

Private Sub Form_AfterUpdate()
    cboID.Requery
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ChangeDate = Now
End Sub

Private Sub Form_Current()
    Dim buildstr As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    AgentNumber.Visible = Me.NewRecord
    cboID.Visible = Not Me.NewRecord
    cboID = AgentNumber
    cboAgentByID = AgentNumber

    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord



        'added 2008-1-2
        'this works for mapquest
        'buildstr = "http://www.mapquest.com/maps/map.adp?formtype="
        'buildstr = buildstr & "&address=" & txtClubAddress1
        'buildstr = buildstr & "&city=" & Trim(.Column(19))
        'buildstr = buildstr & "&state=" & Trim(.Column(20))
        'buildstr = buildstr & "&zipcode=" & Trim(.Column(21))
        If Not IsNull(Trim(Address1)) Or Trim(Address1) <> "" Then
            buildstr = "http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q="
            buildstr = buildstr & Trim(Address1) & ", " & Trim(City) & ", " & Trim(St) & " " & Trim(Zip)
            buildstr = buildstr & "&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
            buildstr = Replace(buildstr, " ", "+")
            lblMapQuest.Caption = "Click here for &Google Map to " & LastName
        Else
            lblMapQuest.Caption = "Click here for &Google Map"
            buildstr = "http://maps.google.com/maps"
        End If


    lblMapQuest.HyperlinkAddress = buildstr
    lblMapQuest.ControlTipText = buildstr



ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Current"
    Resume ExitSub



End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Dirty"
    Resume ExitSub
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)


    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Case 3146 'description = blank
        Me.Undo
        DoCmd.GoToRecord , , acPrevious
        Response = acDataErrContinue

    Case Else
        'MsgBox "Unhandled Form Error " & DataErr & " " & Err.Number & " " & Err.Description
    End Select
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Error"
    Resume ExitSub

End Sub




Private Sub Form_Load()
    SetupDate.ColumnHidden = True
    ChangeDate.ColumnHidden = True
    OrderCount.ColumnHidden = True
    Me.InsideHeight = boxScreen.Height + FormFooter.Height + 100
    Me.InsideWidth = boxScreen.Width + 500

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        If Me.OpenArgs <> "" Then
            Me.LastName.SetFocus
            Me.LastName.Text = Me.OpenArgs
        End If
    End If
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Open"
    Resume ExitSub
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub KMDScustomer_AfterUpdate()
    If KMDScustomer Then
        EMail = "Agent" & AgentNumber & "@kmds.us"
    Else
    End If

End Sub
'------------------------------------------------------------
' cmdfrmDeleteRecord_Click
'
'------------------------------------------------------------
'Private Sub cmdfrmDeleteRecord_Click()
'On Error GoTo cmdfrmDeleteRecord_Click_Err
'
'    On Error Resume Next
'    DoCmd.GoToControl Screen.PreviousControl.Name
'    Err.Clear
'    If (Not Form.NewRecord) Then
'        DoCmd.RunCommand acCmdDeleteRecord
'    End If
'    If (Form.NewRecord And Not Form.Dirty) Then
'        Beep
'    End If
'    If (Form.NewRecord And Form.Dirty) Then
'        DoCmd.RunCommand acCmdUndo
'    End If
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'    End If
'    Call SetRunMailChangesTrue
'
'
'cmdfrmDeleteRecord_Click_Exit:
'    Exit Sub
'
'cmdfrmDeleteRecord_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmDeleteRecord_Click_Exit
'
'End Sub


'------------------------------------------------------------
' cmdfrmSaveRecord_Click
'
''------------------------------------------------------------
'Private Sub cmdfrmSaveRecord_Click()
'On Error GoTo cmdfrmSaveRecord_Click_Err
'
'    On Error Resume Next
'    DoCmd.RunCommand acCmdSaveRecord
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'            End If
'
'
'cmdfrmSaveRecord_Click_Exit:
'    Exit Sub
'
'cmdfrmSaveRecord_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmSaveRecord_Click_Exit
'
'End Sub


'------------------------------------------------------------
' cmdfrmUndo_Click
'
'------------------------------------------------------------
'Private Sub cmdfrmUndo_Click()
'On Error GoTo cmdfrmUndo_Click_Err
'
'    On Error Resume Next
'    DoCmd.RunCommand acCmdUndo
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'            End If
'
'
'cmdfrmUndo_Click_Exit:
'    Exit Sub
'
'cmdfrmUndo_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmUndo_Click_Exit
'
'End Sub


'------------------------------------------------------------
' Command0_Click
'
'------------------------------------------------------------
Private Sub Command0_Click()
On Error GoTo Command0_Click_Err

    DoCmd.Close , ""


Command0_Click_Exit:
    Exit Sub

Command0_Click_Err:
    MsgBox Error$
    Resume Command0_Click_Exit

End Sub


'------------------------------------------------------------
' cmdfrmAddNew_Click
'
'------------------------------------------------------------
'Private Sub cmdfrmAddNew_Click()
'On Error GoTo cmdfrmAddNew_Click_Err
'
'    On Error Resume Next
'    DoCmd.GoToRecord , "", acNewRec
'    If (MacroError <> 0) Then
'        Beep
'        MsgBox MacroError.Description, vbOKOnly, ""
'            End If
'
'
'cmdfrmAddNew_Click_Exit:
'    Exit Sub
'
'cmdfrmAddNew_Click_Err:
'    MsgBox Error$
'    Resume cmdfrmAddNew_Click_Exit
'
'End Sub


'------------------------------------------------------------
' cmdPrintPreview_Click
'
'------------------------------------------------------------
Private Sub cmdPrintPreview_Click()
On Error GoTo cmdPrintPreview_Click_Err

    Call PrintPreView(Form)


cmdPrintPreview_Click_Exit:
    Exit Sub

cmdPrintPreview_Click_Err:
    MsgBox Error$
    Resume cmdPrintPreview_Click_Exit

End Sub


'------------------------------------------------------------
' cmdDataSheet_Click
'
'------------------------------------------------------------
Private Sub cmdDataSheet_Click()
On Error GoTo cmdDataSheet_Click_Err

    Call DatasheetView(Form)


cmdDataSheet_Click_Exit:
    Exit Sub

cmdDataSheet_Click_Err:
    MsgBox Error$
    Resume cmdDataSheet_Click_Exit

End Sub


'------------------------------------------------------------
' cmdAgentsFromOLCCAgents_Click
'
'------------------------------------------------------------
Private Sub cmdAgentsFromOLCCAgents_Click()
On Error GoTo cmdAgentsFromOLCCAgents_Click_Err
    SetOptions False
    DoCmd.OpenQuery "qryAgentsFromOLCCAgentsAppend", acViewNormal, acEdit
    DoCmd.OpenQuery "qryAgentsFromOLCCAgentsUpdate", acViewNormal, acEdit


cmdAgentsFromOLCCAgents_Click_Exit:
    SetOptions True
    MsgBox "Agents have been refreshed... Please Reload Program"
    DoCmd.Close
    Exit Sub

cmdAgentsFromOLCCAgents_Click_Err:
    MsgBox Error$
    Resume cmdAgentsFromOLCCAgents_Click_Exit

End Sub
