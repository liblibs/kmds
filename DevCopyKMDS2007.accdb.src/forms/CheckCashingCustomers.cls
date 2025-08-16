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
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Me![cboID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cboID_AfterUpdate"
    Resume ExitSub

End Sub



Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdfrmAddNew

    DoCmd.GoToRecord , , acNewRec
    FirstName.SetFocus
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

    DoCmd.DoMenuItem acFormBar, acEditMenu, acSelectRecord, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, acDelete, , acMenuVer70



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
If gcfHandleErrors Then On Error GoTo ErrLog

    cmdfrmAddNew.Visible = True
    cmdfrmAddNew.SetFocus
    cmdfrmUndo.Visible = False
    cmdfrmSaveRecord.Visible = False

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub
End Sub



Private Sub Form_AfterUpdate()
    cboID.Requery
    gbRunMailChanges = True

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    'ChangeDate = Now
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error GoTo ErrLog

    'ID.Visible = Me.NewRecord
    cboID.Visible = Not Me.NewRecord
    cboID = ID


    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = Me.NewRecord
    cmdfrmUndo.Visible = Me.NewRecord

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

Private Sub Form_Error(DataErr As Integer, Response As Integer)
Dim nCode As Long
Dim rs As Object

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
    MsgBox Err.Number & " " & Err.Description, , "Sizes:Form_Error"
    Resume ExitSub

End Sub




Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)
End Sub

Private Sub Form_Open(Cancel As Integer)
    gbRunMailChanges = False
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        If Me.OpenArgs <> "" Then
            Me.FirstName.SetFocus
            Me.FirstName.Text = Me.OpenArgs
        Else
            cboID.Visible = True
            'cboID.SetFocus
        End If
        cboID.ColumnHidden = True
    End If
    Me.InsideHeight = boxScreen.Height + FormFooter.Height + 100
    Me.InsideWidth = boxScreen.Width + 500

End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub



Private Sub cmdVendorPOHistory_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdVendorPOHistory_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "VendorPOHistory"

    stLinkCriteria = "[ID]=" & Me![ID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdVendorPOHistory_Click:
    Exit Sub

Err_cmdVendorPOHistory_Click:
    MsgBox Err.Description
    Resume Exit_cmdVendorPOHistory_Click

End Sub
Private Sub Command45_Click()
If gcfHandleErrors Then On Error GoTo Err_Command45_Click

    Dim stDialStr As String
    Dim PrevCtl As Control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91
    Const ERR_CANTMOVE = 2483

    Set PrevCtl = Screen.PreviousControl

    If TypeOf PrevCtl Is TextBox Then
      stDialStr = IIf(varType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ListBox Then
      stDialStr = IIf(varType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ComboBox Then
      stDialStr = IIf(varType(PrevCtl) > V_NULL, PrevCtl, "")
    Else
      stDialStr = ""
    End If

    Application.Run "utility.wlib_AutoDial", stDialStr

Exit_Command45_Click:
    Exit Sub

Err_Command45_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Or (Err = ERR_CANTMOVE) Then
      Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_Command45_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

End Sub
