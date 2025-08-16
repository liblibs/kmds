Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
    gbUpdate = True
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    LineID = Me.Recordset.RecordCount + 1
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Current()
If gcfHandleErrors Then On Error GoTo ErrLog

    LineID.ColumnWidth = -2
    ItemID.ColumnWidth = -2
    ItemDesc.ColumnWidth = -2
    Par.ColumnWidth = -2
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Dim nCode As Double '07/17/2021
    Dim rs As Object
    Case 3022 'duplicate records
        nCode = ItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
        End If
        Response = acDataErrContinue
    Case 3146 'duplicate records
        nCode = ItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
        End If
        Response = acDataErrContinue
    Case 2237
        MsgBox "Please select an item from the list"
        Response = acDataErrContinue

    Case Else
        'MsgBox "Unhandled Form Error " & DataErr & " " & Err.Number & " " & Err.Description
    End Select
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "ClubOrderSub:Form_Error"
    Resume ExitSub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    If Not gcfHandleErrors Then Debug.Print KeyCode
    Select Case KeyCode
    Case 45 'insert
        cmdInsertLine_Click
        KeyCode = 0
    Case 46 'delete
        cmdDelete_Click
        KeyCode = 0
    Case Else
    End Select
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_KeyDown"
    Resume ExitSub

End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
Dim strSQL As String
Dim strSQLWhere As String
    LineID.ColumnOrder = 1
    ItemID.ColumnOrder = 2
    ItemDesc.ColumnOrder = 3
    Par.ColumnOrder = 4

    LineID.ColumnWidth = -2
    ItemID.ColumnWidth = -2
    ItemDesc.ColumnWidth = -2
    Par.ColumnWidth = -2

'    If Me.Parent("chkNonLiquor") Then
'        QuotedPrice.Visible = True
'        strSQLWhere = strSQLWhere & "WHERE (((Item.SalesDepartmentID) <> 1)) "
'
'    Else
'        QuotedPrice.Visible = False
'        strSQLWhere = strSQLWhere & "WHERE (((Item.SalesDepartmentID) = 1)) "
'    End If
'
'    strSQL = "SELECT Item.ID, "
'    strSQL = strSQL & "Item.Description, "
'    strSQL = strSQL & "Price, "
'    strSQL = strSQL & "Size.Description AS SizeD, "
'    strSQL = strSQL & "Item.SalesDepartmentID "
'    strSQL = strSQL & "FROM Item "
'    strSQL = strSQL & " LEFT JOIN [Size] ON Item.SizeID = Size.ID "
'    strSQL = strSQL & strSQLWhere
'    strSQL = strSQL & "ORDER BY Item.ID"
'
'    ItemID.RowSource = strSQL
'
'    strSQL = "SELECT Item.Description, "
'    strSQL = strSQL & "Item.ID, "
'    strSQL = strSQL & "Price, "
'    strSQL = strSQL & "Size.Description AS SizeD, "
'    strSQL = strSQL & "Item.SalesDepartmentID "
'    strSQL = strSQL & "FROM Item "
'    strSQL = strSQL & "   LEFT JOIN [Size] ON Item.SizeID=Size.ID "
'    strSQL = strSQL & strSQLWhere
'    strSQL = strSQL & "ORDER BY Item.Description"
'    ItemDesc.RowSource = strSQL
'


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub cmdDelete_Click()

If gcfHandleErrors Then On Error GoTo Err_cmdDelete_Click
Dim stDocName As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nLineID
Dim rs As Object

    nLineID = Me.Form!LineID

    'add 1 to all lines above
    stDocName = "qryClubOrderGuideTempDeleteLineID"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = stDocName
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = Me.Parent.ID
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters("parLineID") = nLineID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    Me.Requery

    'renumber the recordset
    Me.Form.OrderBy = "LineID"
    Me.Form.OrderByOn = True

    Set rs = Me.Form.RecordsetClone
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

    Me.Requery
    Set rs = Me.Form.RecordsetClone
    rs.FindFirst "[LineID] = " & Str(nLineID)
    If Not rs.EOF Then Me.Form.Bookmark = rs.Bookmark
    rs.Close
    Set rs = Nothing

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
End Sub

Private Sub cmdInsertLine_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdInsertLine_Click

Dim stDocName As String

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nLineID
Dim rs As Object

    nLineID = Me.Form!LineID

    'add 1 to all lines above
    stDocName = "qryClubOrderGuideTempIncrementLineID"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = stDocName
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = Me.Parent.ID
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
        .Parameters("parClubID") = Me.Parent.ID
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters("parLineID") = nLineID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

        'renumber the recordset
    Me.Form.OrderBy = "LineID"
    Me.Form.OrderByOn = True
    Me.Requery

    Set rs = Me.Form.RecordsetClone
    rs.FindFirst "[LineID] = " & Str(nLineID)
    If Not rs.EOF Then Me.Form.Bookmark = rs.Bookmark
    rs.Close
    Set rs = Nothing

Exit_cmdInsertLine_Click:
    ItemID.SetFocus
    Exit Sub

Err_cmdInsertLine_Click:
    MsgBox Err.Description
    Resume Exit_cmdInsertLine_Click

End Sub
