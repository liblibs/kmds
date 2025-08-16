Attribute VB_Name = "modAuditInventory"
Option Compare Database
Option Explicit

Public Sub AuditChange(FieldName As String, ItemID As Double, OldValue As Long, NewValue As Long, Price As Currency)



'Dim rsCashsys As New ADODB.Recordset
Dim rsCashsys As Object
Set rsCashsys = CreateObject("ADODB.Recordset")

'Dim rsInvoLine As New ADODB.Recordset
Dim rsInvoLine As Object
Set rsInvoLine = CreateObject("ADODB.Recordset")

'Dim rsItemMovement As New ADODB.Recordset
Dim rsItemMovement As Object
Set rsItemMovement = CreateObject("ADODB.Recordset")

'Dim rsSalePrice As New ADODB.Recordset
Dim rsSalePrice As Object
Set rsSalePrice = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    Dim nLastInvoiceNo As Long
    Dim nShiftCounter As Long
    Dim strSQL As String
    Dim CurPrice As Currency

    If gcfHandleErrors Then On Error GoTo ErrLog

    '1. get the next invoice #
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsCashsys.BOF And rsCashsys.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
        Exit Sub
    End If

    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 1000000
    nShiftCounter = rsCashsys!ShiftCounter
    rsCashsys.Update
    rsCashsys.Close
    Set rsCashsys = Nothing

    'Is this item on sales
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryItemsSalePriceNowparItemID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("parItemID", adBigInt, adParamInput)
        .Parameters("parItemID") = ItemID
        Set rsSalePrice = .Execute
        If rsSalePrice.BOF And rsSalePrice.EOF Then
            CurPrice = Price
        Else
            rsSalePrice.MoveFirst
            CurPrice = rsSalePrice!SalePrice
        End If

        rsSalePrice.Close
        Set rsSalePrice = Nothing
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")



    '2. update history invoice lines file

    rsInvoLine.Open "InvoLine", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable

    With rsInvoLine
        .AddNew
        !InvoHeadID = nLastInvoiceNo
        !ID = 1 'nLineCount
        !CashSysID = glnCashsysID
        !CashierID = glnCashierID
        !ShiftID = nShiftCounter
        !Code = ItemID
        !InputCode = gstrAdjustReason
        'rsInvoLine!SalesDept = rsItem!SalesDepartmentID
        !Quantity = NewValue - OldValue '!QuantityShipped
        'changed 12/28/2004
        !Price = CurPrice * (NewValue - OldValue)
        !SalesGroup = NewValue
        !TransType = "A" 'Purchase
        !SetupDate = Now
        .Update
        .Close
    End With
    Set rsInvoLine = Nothing

    '3. update the movement file
    strSQL = "SELECT Code, "
    strSQL = strSQL & "Description, "
    strSQL = strSQL & "CurrMM, "
    strSQL = strSQL & "YTDQty "
    strSQL = strSQL & "FROM ItemMovement "
    strSQL = strSQL & "WHERE (((Code)=" & CStr(ItemID) & " ) "
    strSQL = strSQL & "AND ((Description)='Variance Adjustment'))"

    rsItemMovement.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    With rsItemMovement
        If .RecordCount = 0 Then
            .AddNew
            !Code = ItemID
            !Description = "Variance Adjustment"
            !CurrMM = NewValue - OldValue
            '!YTDQty = NewValue - OldValue
        Else
            .MoveFirst
            !CurrMM = Nz(!CurrMM, 0) + (NewValue - OldValue)
            '!YTDQty = Nz(!YTDQty, 0) + (NewValue - OldValue)
        End If
        .Update
        .Close
    End With
    Set rsItemMovement = Nothing

ExitSub:
    Exit Sub
ErrLog:
    fMsgBox "AuditChange" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub


Public Sub AuditPriceChange(FieldName As String, ItemID As Double, OldValue As Currency, NewValue As Currency, Onhand As Long)


'Dim rsCashsys As New ADODB.Recordset
Dim rsCashsys As Object
Set rsCashsys = CreateObject("ADODB.Recordset")

'Dim rsInvoLine As New ADODB.Recordset
Dim rsInvoLine As Object
Set rsInvoLine = CreateObject("ADODB.Recordset")

'Dim rsItemMovement As New ADODB.Recordset
Dim rsItemMovement As Object
Set rsItemMovement = CreateObject("ADODB.Recordset")

'Dim rsSalePrice As New ADODB.Recordset
Dim rsSalePrice As Object
Set rsSalePrice = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    Dim nLastInvoiceNo As Long
    Dim nShiftCounter As Long
    Dim strSQL As String
    Dim CurPrice As Currency

    If gcfHandleErrors Then On Error GoTo ErrLog

    '1. get the next invoice #
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsCashsys.BOF And rsCashsys.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
        Exit Sub
    End If

    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 1000000
    nShiftCounter = rsCashsys!ShiftCounter
    rsCashsys.Update
    rsCashsys.Close
    Set rsCashsys = Nothing


    'Is this item on sales
'    With comQDF
'        .ActiveConnection = CurrentProject.Connection
'        .CommandText = "qryItemsSalePriceNowparItemID"
'        .CommandType = adCmdStoredProc
'        .Parameters.Append .CreateParameter("parItemID", adBigInt, adParamInput)
'        .Parameters("parItemID") = ItemID
'        Set rsSalePrice = .Execute
'        If rsSalePrice.BOF And rsSalePrice.EOF Then
'            CurPrice = Price
'        Else
'            rsSalePrice.MoveFirst
'            CurPrice = rsSalePrice!SalePrice
'        End If
'
'        rsSalePrice.Close
'        Set rsSalePrice = Nothing
'    End With
'    Set comQDF  = Nothing



    '2. update history invoice lines file

    rsInvoLine.Open "InvoLine", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable

    With rsInvoLine
        .AddNew
        !InvoHeadID = nLastInvoiceNo
        !ID = 1 'nLineCount
        !CashSysID = glnCashsysID
        !CashierID = glnCashierID
        !ShiftID = nShiftCounter
        !Code = ItemID
        !InputCode = CStr(OldValue & " to " & NewValue)  'gstrAdjustReason
        !SalesDept = Onhand 'rsItem!SalesDepartmentID
        !Quantity = 0 'Onhand '!QuantityShipped
        !Price = NewValue
        !SalesGroup = (OldValue * 100)
        !TransType = "R" 'Retail price change
        !SetupDate = Now
        .Update
        .Close
    End With
    Set rsInvoLine = Nothing

'    '3. update the movement file
'    strSQL = "SELECT Code, "
'    strSQL = strSQL & "Description, "
'    strSQL = strSQL & "CurrMM, "
'    strSQL = strSQL & "YTDQty "
'    strSQL = strSQL & "FROM ItemMovement "
'    strSQL = strSQL & "WHERE (((Code)=" & CStr(ItemID) & " ) "
'    strSQL = strSQL & "AND ((Description)='Adjustments'))"
'
'    rsItemMovement.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'
'
'    With rsItemMovement
'        If .RecordCount = 0 Then
'            .AddNew
'            !Code = ItemID
'            !Description = "Adjustments"
'            !CurrMM = NewValue - OldValue
'            !YTDQty = NewValue - OldValue
'        Else
'            .MoveFirst
'            !CurrMM = !CurrMM + (NewValue - OldValue)
'            !YTDQty = !YTDQty + (NewValue - OldValue)
'        End If
'        .Update
'        .Close
'    End With
'    Set rsItemMovement = Nothing

ExitSub:
    Exit Sub
ErrLog:
    fMsgBox "AuditPriceChange" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub



Public Function AuditBatchPost() As Boolean

'Dim rsAudit As New ADODB.Recordset
Dim rsAudit As Object
Set rsAudit = CreateObject("ADODB.Recordset")

'Dim rsCashsys As New ADODB.Recordset
Dim rsCashsys As Object
Set rsCashsys = CreateObject("ADODB.Recordset")

'Dim rsInvoLine As New ADODB.Recordset
Dim rsInvoLine As Object
Set rsInvoLine = CreateObject("ADODB.Recordset")

'Dim rsItemMovement As New ADODB.Recordset
Dim rsItemMovement As Object
Set rsItemMovement = CreateObject("ADODB.Recordset")

'Dim rsSalePrice As New ADODB.Recordset
Dim rsSalePrice As Object
Set rsSalePrice = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    Dim nLastInvoiceNo As Long
    Dim nShiftCounter As Long
    Dim strSQL As String
    Dim CurPrice As Currency
    Dim nLineCount As Integer
    Dim datCur As Date

    datCur = Now

    If gcfHandleErrors Then On Error GoTo ErrLog

    '1. get the next invoice #
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsCashsys.BOF And rsCashsys.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
        Exit Function
    End If

    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 1000000
    nShiftCounter = rsCashsys!ShiftCounter
    rsCashsys.Update
    rsCashsys.Close
    Set rsCashsys = Nothing


    '2. update history invoice lines file
    rsAudit.Open "Audit", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable

    rsInvoLine.Open "InvoLine", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable

    Do While Not rsAudit.EOF
        If rsAudit!VarianceAmount <> 0 Then
            nLineCount = nLineCount + 1
            With rsInvoLine
                If Not gcfHandleErrors Then Debug.Print rsAudit!ItemID
                .AddNew
                !InvoHeadID = nLastInvoiceNo
                !ID = nLineCount
                !CashSysID = glnCashsysID
                !CashierID = glnCashierID
                !ShiftID = nShiftCounter
                !Code = rsAudit!ItemID
                !InputCode = "Audit"
                'rsInvoLine!SalesDept = rsItem!SalesDepartmentID
                !Quantity = rsAudit!VarianceAmount
                'changed 12/28/2004
                !Price = rsAudit!Price * rsAudit!VarianceAmount
                !SalesGroup = rsAudit!PhysicalCount
                !TransType = "A" 'Purchase
                !SetupDate = datCur ' same date/time for entire batch...
                .Update
            End With
            strSQL = "UPDATE Item SET Item.Onhand = (Item.Onhand + " & rsAudit!VarianceAmount & ") "
            strSQL = strSQL & "WHERE (((Item.ID)= " & CStr(rsAudit!ItemID) & "))"
            With comQDF
                'qryPurchaseOrderItemReplenishMake
                .ActiveConnection = CurrentProject.Connection
                .CommandText = strSQL
                .CommandType = adCmdText
                .Execute
            End With
            Set comQDF = Nothing
            Set comQDF = CreateObject("ADODB.COMMAND")

            rsAudit!VarianceAmount = 0
            rsAudit.Update
        End If

        rsAudit.MoveNext

    Loop

    rsInvoLine.Close
    Set rsInvoLine = Nothing
    Set rsItemMovement = Nothing
    rsAudit.Close
    Set rsAudit = Nothing

    strSQL = "Delete from Audit Where (VarianceAmount = 0)"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If nLineCount > 0 Then
        SQLPassThru "sp_UpdateItemMovement"
        'UpdateMovement (False)
    End If
    AuditBatchPost = True
ExitSub:
    Exit Function
ErrLog:
    fMsgBox "AuditChange" & vbCrLf & Err.Number & " " & Err.Description
    AuditBatchPost = False
    Resume ExitSub
End Function

Public Sub AuditFileAppendByPDA()


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    Dim nLastInvoiceNo As Long
    Dim nShiftCounter As Long
    Dim strSQL As String
    Dim CurPrice As Currency
    Dim nLineCount As Integer
    Dim datCur As Date

    datCur = Now

    If gcfHandleErrors Then On Error GoTo ErrLog

    'delete file?
    strSQL = "Delete from Audit"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    strSQL = "qryPSCAuditAppend"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Execute
    End With
ExitSub:
    Set comQDF = Nothing
    Exit Sub
ErrLog:
    fMsgBox "AuditFileAppendByPDA" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub
