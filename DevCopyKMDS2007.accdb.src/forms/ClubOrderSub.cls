Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim mbDeleteRecord As Boolean
Dim bFintechLiquorOnly As Boolean

Private Sub Form_Activate()

    LineID.ColumnWidth = -2
    ItemID.ColumnWidth = -2
    Description.ColumnWidth = -2
    Size.ColumnWidth = 1000
    Category.ColumnWidth = -2
    Quantity.ColumnWidth = -2
    Price.ColumnWidth = -2
    SalePrice.ColumnWidth = -2
    Cost.ColumnWidth = -2
    ExtPrice.ColumnWidth = -2
    ExtCost.ColumnWidth = -2
    Me.DatasheetBackColor = LineID.BackColor

End Sub

Private Sub Form_AfterUpdate()
Dim rs As Recordset
    ItemID.ColumnWidth = -2
    Description.ColumnWidth = -2


    ClubOrderRecalcTotals (OrdheadID)
'
'    If Me.Recordset.RecordCount = 1 Then
'        With Me.Parent
'            .cmbOrderID.Requery
'            .txtDiscountRate = .cmbOrderID.Column(13)
'            .txtNetPrice = .cmbOrderID.Column(15)
'        End With
'        'reformat the columns for best fit
'        Form_Load
'
'
'
'    End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nNextLine As Integer

If gcfHandleErrors Then On Error GoTo ErrLog


     'delete original
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryOrderLineMaxLineID"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parOrdheadID", adBigInt, adParamInput)
        .Parameters("parOrdheadID") = OrdheadID
        Set rstemp = .Execute
        rstemp.MoveFirst
        If IsNull(rstemp!MaxLineID) Then
            nNextLine = 1
        Else
            nNextLine = rstemp!MaxLineID + 1
        End If
        rstemp.Close
        Set rstemp = Nothing
    End With
    Set comQDF = Nothing


    LineID = nNextLine 'Me.Recordset.RecordCount + 1
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "ClubOrderSub:Form_BeforeInsert"
    Resume ExitSub
End Sub



Private Sub Form_BeforeUpdate(Cancel As Integer)
    'check for availability
    Dim strSQL As String
    If Me.Parent.Label50.Caption = "Fintech" Then
        bFintechLiquorOnly = GetKMDSSettings("FintechLiquorOnly", "True")

        If bFintechLiquorOnly And ItemID.Column(11) <> 1 Then
            MsgBox "Fintech Liquor ONLY!"
            Cancel = True
            Me.Undo
            Exit Sub
        End If
    End If

    If IsNull(ItemID) Then
        Description.SetFocus
        Cancel = True
        Exit Sub
    End If
    If Quantity > 0 And Quantity > CLng(Nz(ItemID.Column(3), 0)) Then
        strSQL = "Warning: Quantity ordered is not on the shelf!" & vbCrLf & vbCrLf
        strSQL = strSQL & ItemID.Column(0) & " " & ItemID.Column(1) & vbCrLf & vbCrLf
        strSQL = strSQL & " Ordered = " & Quantity & vbCrLf & vbCrLf
        strSQL = strSQL & "On Shelf = " & ItemID.Column(3) & vbCrLf & vbCrLf
        strSQL = strSQL & "Reserved on other Club Orders = " & ItemID.Column(9) & vbCrLf
        strSQL = strSQL & "Purchase Order = " & ItemID.Column(10) & vbCrLf


        If MsgBox(strSQL, vbOKCancel, "Item ordered is Not available") = vbCancel Then
            Cancel = True
            Quantity.SetFocus
        End If
    End If
End Sub

Private Sub Form_Current()
    'just stop at quantities for existing records
    If gcfHandleErrors Then On Error Resume Next
    Price.ColumnWidth = -2

    ItemID.TabStop = Me.NewRecord
    Price.Locked = Not Nz(Me.Parent.chkNonLiquor, 0)
    'added 2007-12-31
    If Me.NewRecord Then
        If Me.Parent("chkSkipQuantityOnNewItems") Then
            ItemID.SetFocus
        End If
    Else
        Quantity.SetFocus
    End If

'    If mbDeleteRecord Then
'        ClubOrderRecalcTotals (OrdheadID)
'        mbDeleteRecord = False
'    End If

'    If Me.NewRecord And Me.Recordset.RecordCount Then
        'FORCE THE EDIT TO BEGIN
'        Quantity = 0
'    End If
End Sub

Private Sub Form_Delete(Cancel As Integer)
'    mbDeleteRecord = True
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case DataErr
    Dim nCode As Double
    Dim rs As Object
    Case 3022 'duplicate records
        nCode = ItemID

        If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nCode, vbYesNo) = vbYes Then
            Me.Form.Undo
            If gcfHandleErrors Then On Error GoTo ErrLog
            Set rs = Me.Recordset.Clone
            rs.FindFirst "[ItemID] = " & Str(nCode)
            Me.Bookmark = rs.Bookmark
            Quantity.SetFocus
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
            Quantity.SetFocus
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
    Select Case KeyCode
    Case 119
        cmdAddRecord_Click
        KeyCode = 0
    Case Else

    End Select
End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog

    LineID.ColumnOrder = 1
    Quantity.ColumnOrder = 2
    ItemID.ColumnOrder = 3
    Description.ColumnOrder = 4
    Size.ColumnOrder = 5
    cmbShelf.ColumnOrder = 6
    Category.ColumnOrder = 7
    SalePrice.ColumnOrder = 8
    Price.ColumnOrder = 9

    Buydown.ColumnOrder = 10
    Cost.ColumnOrder = 11
    ExtPrice.ColumnOrder = 12
    ExtCost.ColumnOrder = 13

    cmbShelf.ColumnWidth = -2

    LineID.ColumnWidth = -2
    ItemID.ColumnWidth = -2
    Description.ColumnWidth = -2
    Size.ColumnWidth = 1000
    Quantity.ColumnWidth = -2
    Price.ColumnWidth = -2
    SalePrice.ColumnWidth = -2
    Buydown.ColumnWidth = -2
    Cost.ColumnWidth = -2
    ExtPrice.ColumnWidth = -2
    ExtCost.ColumnWidth = -2
    Me.DatasheetBackColor = LineID.BackColor
 '   mbDeleteRecord = False


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub

End Sub

Private Sub Description_AfterUpdate()
    ItemID_AfterUpdate
End Sub

Private Sub Form_Open(Cancel As Integer)

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Open"
    Resume ExitSub

End Sub

Private Sub ItemID_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
'    'add the buydown to the price; it is discounted at the register
'    Price = CCur((ItemID.Column(5)) + CCur(ItemID.Column(13)))
    'NOT ANYMORE!
    Price = CCur(ItemID.Column(5))
    SalePrice = ItemID.Column(6)
    Buydown = CCur(ItemID.Column(13))
    'added 2007-12-31
    If Me.NewRecord Then
        If Me.Parent("chkSkipQuantityOnNewItems") Then
            If IsNull(Quantity) Or Quantity = 0 Then
                Quantity = 1
            End If
        End If
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ItemID_AfterUpdate"
    Resume ExitSub
End Sub



Private Sub ItemID_DblClick(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo Err_cmdOpenItem_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Item"

    stLinkCriteria = "[ID]=" & Me![ItemID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpenItem_Click:
    Exit Sub

Err_cmdOpenItem_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenItem_Click
End Sub




Private Sub ItemID_Exit(Cancel As Integer)
'    MsgBox "exit"
End Sub

Private Sub LineID_DblClick(Cancel As Integer)
    If MsgBox("Do you want to insert an item at line " & LineID, vbYesNo) = vbYes Then
        InsertLine (LineID)
    End If
End Sub

Private Sub Size_BeforeUpdate(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog

    MsgBox "Change size by picking a different Item from the ItemID column."
    Size.Undo
    Cancel = True
'    ItemID.SetFocus
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Size_BeforeUpdate"
    Resume ExitSub
End Sub

Private Sub InsertLine(AtLineID As Integer)

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nItemCode As Double
Dim rs
Dim rsLookup

    nItemCode = InputBox("Please enter the code of the item to insert...", "Insert item at line #" & AtLineID)
    '1. is this a valid item code?
    '2. is this a duplicate item?
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryOrderLineItemID"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parOrdheadID", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("parItemID", adDecimal, adParamInput)
        .Parameters("parOrdheadID") = OrdheadID
        .Parameters("parItemID") = nItemCode
        Set rstemp = .Execute
    End With

    With rstemp
        If Not .BOF And Not .EOF Then
            .MoveFirst
            If MsgBox("Duplicate records not allowed; Do you want to go to line with code =" & nItemCode, vbYesNo) = vbYes Then
                If gcfHandleErrors Then On Error GoTo ErrLog
                Set rs = Me.Recordset.Clone
                rs.FindFirst "[ItemID] = " & Str(nItemCode)
                Me.Bookmark = rs.Bookmark
                rs.Close
                Set rs = Nothing
            End If
            Exit Sub
        End If
        .Close
    End With

    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")

    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    Me.Parent.lblRetr.Visible = True
    Me.Parent.Form.Repaint

    '3. push up all the lines greater than AtLineID
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryOrderLineInsertLineNumbers"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parOrdheadID", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters("parOrdheadID") = OrdheadID
        .Parameters("parLineID") = AtLineID
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    '4. Insert the new item atlineid

'    Set RS = Me.Recordset.Clone
'    RS.AddNew
'    RS!LineID = AtLineID
'    RS!OrdheadID = OrdheadID
'    RS!ItemID = nItemCode
'
'    RS.Update
'    RS.Close
'
'    Set RS = Nothing

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryOrderLineInsertNewItem"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parOrdheadID", adBigInt, adParamInput)
        .Parameters.Append .CreateParameter("parLineID", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("parItemID", adDouble, adParamInput)
        .Parameters("parOrdheadID") = OrdheadID
        .Parameters("parLineID") = AtLineID
        .Parameters("parItemID") = CDbl(nItemCode)
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    Me.Refresh
    Me.OrderByOn = False
    Me.OrderByOn = True

    Set rsLookup = Me.Recordset.Clone
    rsLookup.FindFirst "[ItemID] = " & Str(nItemCode)
    Me.Bookmark = rsLookup.Bookmark
    rsLookup.Close
    Set rsLookup = Nothing
    ItemID_AfterUpdate


ExitSub:
    Me.Parent.lblRetr.Visible = False
    Me.Parent.Form.Repaint

    Exit Sub

ErrLog:
    EMsgBox Me, "InsertLine"
    Resume ExitSub
End Sub
Private Sub cmdAddNewRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddNewRecord_Click


    DoCmd.GoToRecord , , acNewRec

Exit_cmdAddNewRecord_Click:
    Exit Sub

Err_cmdAddNewRecord_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNewRecord_Click

End Sub
Private Sub cmdAddRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddRecord_Click


    DoCmd.GoToRecord , , acNewRec
    ItemID.SetFocus
Exit_cmdAddRecord_Click:
    Exit Sub

Err_cmdAddRecord_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddRecord_Click

End Sub
