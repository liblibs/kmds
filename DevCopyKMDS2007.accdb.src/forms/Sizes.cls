Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
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


Private Sub Command29_Click()
ExportReport Me.Tag

End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    cboID.Requery

End Sub

Private Sub Form_AfterInsert()
    cboID.Requery

End Sub

Private Sub Form_AfterUpdate()
    cboID.Requery

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ChangeDate = Now
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me

End Sub

Private Sub Form_Current()

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    If gcfHandleErrors Then On Error GoTo ErrLog
    If Me.NewRecord Then


        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandType = adCmdStoredProc
            .CommandText = "qrySizeMaxID" 'strSQL"
        End With

        Set rstemp = comQDF.Execute

        ID.Visible = True
        ID = rstemp!maxofid + 1

        rstemp.Close
        Set comQDF = Nothing
        Set rstemp = Nothing
        cboID.Visible = False
        [Desc].SetFocus
    Else
        cboID.Visible = True
        ID.Visible = False
        cboID = ID

    End If

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

Private Sub Form_DblClick(Cancel As Integer)
    cboID.ColumnHidden = Not cboID.ColumnHidden
    SetupDate.ColumnHidden = Not SetupDate.ColumnHidden
    ChangeDate.ColumnHidden = Not ChangeDate.ColumnHidden

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



Private Sub Form_Load()
    cboID.ColumnHidden = True
    SetupDate.ColumnHidden = True
    ChangeDate.ColumnHidden = True
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        If Me.OpenArgs <> "" Then
            Me.Desc.SetFocus
            Me.Desc.Text = Me.OpenArgs
        End If
    End If
End Sub

Private Sub ID_AfterUpdate()
    If optNewLiquorSize Then
        If ID > 0 And ID < 100 Then
        Else
            MsgBox "Size must be between 1 and 99 for liquor items"
        End If

    Else
        If ID > 0 And ID < 100 Then
            MsgBox "Size must be greater than 99 for non-liquor items"
        Else
        End If
    End If

End Sub

Private Sub optNewLiquorSize_AfterUpdate()
    If optNewLiquorSize Then
        DoCmd.RunMacro "macKMDSAdd"
        ID.Visible = True
        ID.Value = 0
        ID.SetFocus
    Else
        Desc.SetFocus
    End If

End Sub

Private Sub Text30_DblClick(Cancel As Integer)
Text30.Enabled = True
End Sub
