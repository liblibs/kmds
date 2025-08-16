Attribute VB_Name = "modPurchaseOrder"
Option Compare Database
Option Explicit

Dim mnLinecount As Integer


Public Sub AppendUPCItemsWithItemID()
'the ID column of the item row must be an ID column in the upc row
' to do this Item!ID must be converted to a string - qryUPCItemsAppendItemID_01_Make_ItemUPCID
' so a join can be made in the append - qryUPCItemsAppendItemID_02_Append




If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    SetOptions False
    If gcfHandleErrors Then On Error Resume Next
    DoCmd.DeleteObject acTable, "ItemUPCID"

    If gcfHandleErrors Then On Error GoTo ErrLog
    stDocName = "qryUPCItemsAppendItemID_01_Make_ItemUPCID"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    stDocName = "qryUPCItemsAppendItemID_02_Append"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
ExitSub:
    SetOptions True
    Exit Sub
ErrLog:
    MsgBox "AppendUPCItemsWithItemID" & vbCrLf & Err.Number & " " & Err.Description
End Sub


Public Function CreatePurchaseOrder(VendorID As Long, IncludeSaleItems As Boolean, IncludeTPR As Boolean) As Integer

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrder = 0
    'ArchiveReport "PurchaseOrderCreate"
    'MsgBox "CreatePurchaseOrder 1"
     With comQDF
        .ActiveConnection = CurrentProject.Connection

        'added 12/27/2004
        If VendorID = 1 Then
            'added 11/20/2006 for bert
            If Forms("PurchaseOrderCreate").optBasePO = 7 Then
                .CommandText = "qryPurchaseOrderCreateWithNoOnOrders"
            Else
                .CommandText = "qryPurchaseOrderCreate"
            End If
        Else
            'same query/ different sort (sales dept/description)
            .CommandText = "qryPurchaseOrderCreateByDeptByDesc"

        End If
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    'MsgBox "CreatePurchaseOrder 2"
    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number
    'MsgBox "CreatePurchaseOrder 3"
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    'MsgBox "CreatePurchaseOrder 4"
    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    'MsgBox "CreatePurchaseOrder 5"
    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With
'
    'MsgBox "CreatePurchaseOrder 6"
    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    'MsgBox "CreatePurchaseOrder 7"
    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ItemID
            rsPOLine!QuantityOrdered = !SOQ
            rsPOLine!QuantityShipped = !SOQ
            rsPOLine!Price = !Sprice
            rsPOLine!Cost = !Cost
            rsPOLine!Repack = !Repack

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SOQ
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")


    'MsgBox "CreatePurchaseOrder 8"
    If IncludeSaleItems = True Then
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")
        'MsgBox "CreatePurchaseOrder 8a"
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderCreateSaleItems"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters("parVendorID") = VendorID
        End With

        'MsgBox "CreatePurchaseOrder 8b"
        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

        With rstemp
            'MsgBox "CreatePurchaseOrder 8c"
            If .RecordCount > 0 Then

                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
                'GoTo ErrLogShowRstemp
                Do While Not .EOF
                    'MsgBox "CreatePurchaseOrder 8d"
                    nLineCount = nLineCount + 1

                    rsPOLine.AddNew
                    rsPOLine!VendorID = VendorID
                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
                    rsPOLine!LineID = nLineCount
                    rsPOLine!ItemID = !ItemID
                    rsPOLine!QuantityOrdered = 0
                    rsPOLine!QuantityShipped = 0
                    rsPOLine!Price = !SalePrice
                    rsPOLine!Cost = !Cost

                    rsPOLine!SetupDate = dtSetup
                    rsPOLine!ChangeDate = dtSetup
                    rsPOLine.Update
                    If Not gcfHandleErrors Then Debug.Print !ItemID
                    .MoveNext
                    'MsgBox "CreatePurchaseOrder 8E"

                Loop
            End If
            .Close
        End With
        Set rstemp = Nothing
        Set rstemp = CreateObject("ADODB.Recordset")
    End If

    'MsgBox "CreatePurchaseOrder 9"
    If IncludeTPR = True Then
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderCreateTPRItems"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters("parVendorID") = VendorID
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

        With rstemp
            If .RecordCount > 0 Then

                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
                'GoTo ErrLogShowRstemp
                Do While Not .EOF
                    nLineCount = nLineCount + 1
                    rsPOLine.AddNew
                    rsPOLine!VendorID = VendorID
                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
                    rsPOLine!LineID = nLineCount
                    rsPOLine!ItemID = !ItemID
                    rsPOLine!QuantityOrdered = 0
                    rsPOLine!QuantityShipped = 0
                    rsPOLine!Price = !Price
                    rsPOLine!Cost = !Cost

                    rsPOLine!SetupDate = dtSetup
                    rsPOLine!ChangeDate = dtSetup
                    rsPOLine.Update
                    If Not gcfHandleErrors Then Debug.Print !ItemID;
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rstemp = Nothing
        Set rstemp = CreateObject("ADODB.Recordset")
    End If


ExitSub:
    CreatePurchaseOrder = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function CreatePurchaseOrderItemsSupplied(VendorID As Long) As Integer

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")



Dim nOrderCount As Integer
Dim nTransferID As Integer
Dim nLineCount As Integer
Dim dtSetup As Date

'MsgBox "Create 1"
If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderItemsSupplied = 0
    'ArchiveReport "PurchaseOrderCreate"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCreateItemsSupplied"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

'    MsgBox "Create 2"
    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

'    MsgBox "Create 3"

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nTransferID = !ID
        .Close
       Set rsPO = Nothing
    End With

'    MsgBox "Create 4"

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

'    MsgBox "Create 5"
    With rstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nTransferID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ItemID
            'added 12/27/2004
            If (!OrderTo > 0) And (!Onhand <= !OrderAt) Then
                rsPOLine!QuantityOrdered = IIf(!SOQ > 0, !SOQ, 0)
                rsPOLine!QuantityShipped = IIf(!SOQ > 0, !SOQ, 0)
            Else
                rsPOLine!QuantityOrdered = 0
                rsPOLine!QuantityShipped = 0
            End If
            rsPOLine!Price = !Sprice
            rsPOLine!Cost = !Cost
            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SOQ
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

ExitSub:
    CreatePurchaseOrderItemsSupplied = nTransferID
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function CreatePurchaseOrderBlankOrder(VendorID As Long) As Integer

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")


Dim nOrderCount As Integer
Dim nTransferID As Integer
Dim nLineCount As Integer
Dim dtSetup As Date

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderBlankOrder = 0

    'get the next PO number
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nTransferID = !ID
        .Close
       Set rsPO = Nothing
    End With


ExitSub:
    CreatePurchaseOrderBlankOrder = nTransferID
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function CreatePurchaseOrderDaysSupply(VendorID As Long, IncludeSaleItems As Boolean, IncludeTPR As Boolean) As Integer

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String
Dim strSQL As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderDaysSupply = 0
    'ArchiveReport "PurchaseOrderCreate"

    'strSQL = "SELECT * FROM PurchaseOrderItemDaysSupply "
    'strSQL = strSQL & "WHERE (((SOQ)>0))"

    strSQL = "SELECT PurchaseOrderItemDaysSupply.* "
    strSQL = strSQL & "FROM (PurchaseOrderItemDaysSupply "
    strSQL = strSQL & "INNER JOIN Item ON PurchaseOrderItemDaysSupply.ItemId = Item.ID) "
    strSQL = strSQL & "INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID "
    strSQL = strSQL & "WHERE (((SubCategory.Description)<>'OBSOLETE') "
    strSQL = strSQL & "AND ((PurchaseOrderItemDaysSupply.SOQ)>0))"

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Set rstemp = CreateObject("ADODB.Recordset")
            Exit Function
        End If
    End With

    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            If !SOQ > 0 Then
                nLineCount = nLineCount + 1
                rsPOLine.AddNew
                rsPOLine!VendorID = VendorID
                rsPOLine!PurchaseOrderID = nPurchaseOrderID
                rsPOLine!LineID = nLineCount
                rsPOLine!ItemID = !ItemID
                rsPOLine!QuantityOrdered = !SOQ
                rsPOLine!QuantityShipped = !SOQ
                rsPOLine!Price = !Sprice
                rsPOLine!Cost = !Cost
                rsPOLine!Repack = !Repack

                rsPOLine!SetupDate = dtSetup
                rsPOLine!ChangeDate = dtSetup
                rsPOLine.Update
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")

    If IncludeSaleItems = True Then
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderCreateSaleItems"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular

'            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
'            .Parameters("parPurchaseOrderID") = nPurchaseOrderID
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters("parVendorID") = VendorID
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

        With rstemp
            If .RecordCount > 0 Then

                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
                'GoTo ErrLogShowRstemp
                Do While Not .EOF
                    nLineCount = nLineCount + 1
                    rsPOLine.AddNew
                    rsPOLine!VendorID = VendorID
                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
                    rsPOLine!LineID = nLineCount
                    rsPOLine!ItemID = !ItemID
                    rsPOLine!QuantityOrdered = 0
                    rsPOLine!QuantityShipped = 0
                    rsPOLine!Price = !SalePrice
                    rsPOLine!Cost = !Cost

                    rsPOLine!SetupDate = dtSetup
                    rsPOLine!ChangeDate = dtSetup
                    rsPOLine.Update
                    If Not gcfHandleErrors Then Debug.Print !Code; " "; !SOQ
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rstemp = Nothing
        Set rstemp = CreateObject("ADODB.Recordset")
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")
    End If

    If IncludeTPR = True Then
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderCreateTPRItems"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular

'            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
'            .Parameters("parPurchaseOrderID") = nPurchaseOrderID
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters("parVendorID") = VendorID
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

        With rstemp
            If .RecordCount > 0 Then

                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
                'GoTo ErrLogShowRstemp
                Do While Not .EOF
                    nLineCount = nLineCount + 1
                    rsPOLine.AddNew
                    rsPOLine!VendorID = VendorID
                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
                    rsPOLine!LineID = nLineCount
                    rsPOLine!ItemID = !ItemID
                    rsPOLine!QuantityOrdered = 0
                    rsPOLine!QuantityShipped = 0
                    rsPOLine!Price = !Price
                    rsPOLine!Cost = !Cost

                    rsPOLine!SetupDate = dtSetup
                    rsPOLine!ChangeDate = dtSetup
                    rsPOLine.Update
                    If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SOQ
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rstemp = Nothing
        Set rstemp = CreateObject("ADODB.Recordset")
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

    End If

ExitSub:
    CreatePurchaseOrderDaysSupply = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!Code
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function CreatePurchaseOrderFromMetroLogic(VendorID As Long) As Integer

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderFromMetroLogic = 0
    'ArchiveReport "PurchaseOrderCreate"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryMetrologicScanBatchForOrders"
        .CommandType = adCmdStoredProc
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ItemID
            rsPOLine!QuantityOrdered = !SumOfQuantity
            rsPOLine!QuantityShipped = !SumOfQuantity
            rsPOLine!Price = !Price
            rsPOLine!Cost = !Cost
            rsPOLine!Repack = IIf(!SumOfQuantity < !UnitsPerCase, True, False)

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SumOfQuantity
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set comQDF = Nothing
        'delete the work file
    SetOptions False
    DoCmd.OpenQuery "qryMetrologicScanBatchDelete"
    SetOptions True

ExitSub:
    CreatePurchaseOrderFromMetroLogic = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function
Public Function CreatePurchaseOrderFromPricingTool(VendorID As Long) As Integer

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderFromPricingTool = 0
    'ArchiveReport "PurchaseOrderCreate"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPricingToolForPurchasOrder"
        .CommandType = adCmdStoredProc
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ID
            rsPOLine!QuantityOrdered = 1
            rsPOLine!QuantityShipped = 1
            rsPOLine!Price = !Price
            rsPOLine!Cost = !Cost
            rsPOLine!Repack = IIf(1 < !UnitsPerCase, True, False)

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            If Not gcfHandleErrors Then Debug.Print !ID;
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set comQDF = Nothing
        'delete the work file

ExitSub:
    CreatePurchaseOrderFromPricingTool = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function


Public Function CreatePurchaseOrderFromScanBatch(VendorID As Long) As Integer

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderFromScanBatch = 0
    'ArchiveReport "PurchaseOrderCreate"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryScanBatchForOrders"
        .CommandType = adCmdStoredProc
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ItemID
            rsPOLine!QuantityOrdered = !SumOfCountQuantity
            rsPOLine!QuantityShipped = !SumOfCountQuantity
            rsPOLine!Price = !FirstOfPrice
            rsPOLine!Cost = !FirstOfCost
            rsPOLine!Repack = IIf(!SumOfCountQuantity >= !UnitsPerCase, True, False)

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SumOfCountQuantity
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

ExitSub:
    CreatePurchaseOrderFromScanBatch = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function


Public Function CreatePurchaseOrderWADS(VendorID As Long, IncludeSaleItems As Boolean, IncludeTPR As Boolean) As Integer

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String
Dim strSQL As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderWADS = 0
    'ArchiveReport "PurchaseOrderCreate"

    strSQL = "SELECT * FROM PurchaseOrderItemWADS "
    strSQL = strSQL & "WHERE ((SOQRaw)>1) "
    'strSQL = strSQL & "OR ( ((PurchaseOrderItemWADS.OrderMax)<[OrderMultiple]*0.5) "
    'strSQL = strSQL & "  AND ((PurchaseOrderItemWADS.SOQRaw)>1))"



    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdText
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ItemID
            If !SOQ > 0 Then
                rsPOLine!QuantityOrdered = !SOQ
                rsPOLine!QuantityShipped = !SOQ
            Else
                rsPOLine!QuantityOrdered = !OrderMultiple
                rsPOLine!QuantityShipped = !OrderMultiple
            End If
            If rsPOLine!QuantityShipped < !UnitsPerCase Then
                rsPOLine!Repack = -1
            Else
                rsPOLine!Repack = 0
            End If
            rsPOLine!Price = !Price
            rsPOLine!Cost = !Cost

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")


    If IncludeSaleItems = True Then
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderCreateSaleItems"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular

'            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
'            .Parameters("parPurchaseOrderID") = nPurchaseOrderID
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters("parVendorID") = VendorID
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

        With rstemp
            If .RecordCount > 0 Then

                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
                'GoTo ErrLogShowRstemp
                Do While Not .EOF
                    nLineCount = nLineCount + 1
                    rsPOLine.AddNew
                    rsPOLine!VendorID = VendorID
                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
                    rsPOLine!LineID = nLineCount
                    rsPOLine!ItemID = !ItemID
                    rsPOLine!QuantityOrdered = 0
                    rsPOLine!QuantityShipped = 0
                    rsPOLine!Price = !SalePrice
                    rsPOLine!Cost = !Cost

                    rsPOLine!SetupDate = dtSetup
                    rsPOLine!ChangeDate = dtSetup
                    rsPOLine.Update
                    If Not gcfHandleErrors Then Debug.Print !Code; " "; !SOQ
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Set rstemp = Nothing
    Set rstemp = CreateObject("ADODB.Recordset")
    End If

    If IncludeTPR = True Then
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderCreateTPRItems"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular

'            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
'            .Parameters("parPurchaseOrderID") = nPurchaseOrderID
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters("parVendorID") = VendorID
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

        With rstemp
            If .RecordCount > 0 Then

                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
                'GoTo ErrLogShowRstemp
                Do While Not .EOF
                    nLineCount = nLineCount + 1
                    rsPOLine.AddNew
                    rsPOLine!VendorID = VendorID
                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
                    rsPOLine!LineID = nLineCount
                    rsPOLine!ItemID = !ItemID
                    rsPOLine!QuantityOrdered = 0
                    rsPOLine!QuantityShipped = 0
                    rsPOLine!Price = !Price
                    rsPOLine!Cost = !Cost

                    rsPOLine!SetupDate = dtSetup
                    rsPOLine!ChangeDate = dtSetup
                    rsPOLine.Update
                    If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SOQ
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rstemp = Nothing
        Set rstemp = CreateObject("ADODB.Recordset")
    End If


ExitSub:
    CreatePurchaseOrderWADS = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function


Public Function CreatePurchaseOrderFromOLCCEPS(VendorID As Long, EPSOrderNumber As Long, EPSOrderDate) As Integer

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim rsVendor As New ADODB.Recordset
Dim rsVendor As Object
Set rsVendor = CreateObject("ADODB.Recordset")

'Dim rsPO As New ADODB.Recordset
Dim rsPO As Object
Set rsPO = CreateObject("ADODB.Recordset")

'Dim rsPOLine As New ADODB.Recordset
Dim rsPOLine As Object
Set rsPOLine = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreatePurchaseOrderFromOLCCEPS = 0
    'ArchiveReport "PurchaseOrderCreate"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryOLCCEPSForOrders"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("prmOrderNumber", adBigInt, adParamInput)
        .Parameters("prmOrderNumber") = EPSOrderNumber

        .Parameters.Append .CreateParameter("prmOrderDate", adDate, adParamInput)
        .Parameters("prmOrderDate") = Format(EPSOrderDate, "mm/dd/yyyy 00:00")
        '.Parameters("BeginDate") = Format(OrderDate, "mm/dd/yyyy 00:00")

    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parVendorID", adInteger, adParamInput)
        .Parameters("parVendorID") = VendorID
    End With

    With rsVendor
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        .Update
        .Close
        Set rsVendor = Nothing
    End With

    With rsPO
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrder", .ActiveConnection, , , adCmdTable
        .AddNew
        !VendorID = VendorID
        !OrderCount = nOrderCount
        '1/2/23 use as indicator of electronic packing list
        !SetupDate = #1/1/2018#

        .Update
        .MoveLast
        nPurchaseOrderID = !ID
        .Close
       Set rsPO = Nothing
    End With

    With rsPOLine
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "PurchaseOrderLine", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !OLCCItem
            rsPOLine!QuantityOrdered = !SumOfShipped
            rsPOLine!QuantityShipped = !SumOfShipped
            rsPOLine!Price = !FirstOfCurPrice
            rsPOLine!Cost = !FirstOfPrice
            rsPOLine!Repack = IIf(!SumOfShipped >= !FirstOfUnitsPerCase, False, True)

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update
            If Not gcfHandleErrors Then Debug.Print !OLCCItem; " "; !Shipped
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

ExitSub:
    CreatePurchaseOrderFromOLCCEPS = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & "CreatePurchaseOrderFromOLCCEPS #" & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function CreateStoreTransferFromScanBatch(AgentID As Long, TransferID As Long, InOrOut As String) As Integer

'Dim rsAgentTransferLines As New ADODB.Recordset
Dim rsAgentTransferLines As Object
Set rsAgentTransferLines = CreateObject("ADODB.Recordset")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreateStoreTransferFromScanBatch = 0
    'ArchiveReport "PurchaseOrderCreate"
    'MsgBox "CreateStoreTransferFromScanBatch  1"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryScanBatchForOrders"
        .CommandType = adCmdStoredProc
    End With

    'MsgBox "CreateStoreTransferFromScanBatch  2"
    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number


    With rsAgentTransferLines
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "AgentTransferLines", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsAgentTransferLines.AddNew

            rsAgentTransferLines!AgentTransferHeadID = TransferID
            rsAgentTransferLines!LineID = nLineCount
            rsAgentTransferLines!ItemID = !ItemID
            rsAgentTransferLines!QuantityShipped = !SumOfCountQuantity
            If InOrOut = "I" Then
                rsAgentTransferLines!QuantityOrdered = !SumOfCountQuantity
                '!Status = "I"
            Else
                rsAgentTransferLines!QuantityOrdered = 0 - !SumOfCountQuantity
                '!Status = "O"
            End If

            rsAgentTransferLines!Price = !FirstOfPrice
            rsAgentTransferLines!Cost = !FirstOfCost

            rsAgentTransferLines!SetupDate = dtSetup
            rsAgentTransferLines!ChangeDate = dtSetup
            rsAgentTransferLines.Update
            If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SumOfCountQuantity
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    CreateStoreTransferFromScanBatch = TransferID

ExitSub:
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function CreateStoreTransferFromMetrologic(AgentID As Long, TransferID As Long, InOrOut As String) As Integer

'Dim rsAgentTransferLines As New ADODB.Recordset
Dim rsAgentTransferLines As Object
Set rsAgentTransferLines = CreateObject("ADODB.Recordset")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreateStoreTransferFromMetrologic = 0
    'ArchiveReport "PurchaseOrderCreate"
    'MsgBox "CreateStoreTransferFromScanBatch  1"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryMetrologicScanBatchForOrders"
        .CommandType = adCmdStoredProc
    End With

    'MsgBox "CreateStoreTransferFromScanBatch  2"
    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly

    With rstemp
        If .RecordCount = 0 Then
            .Close
            Set rstemp = Nothing
            Exit Function
        End If
    End With

    'get the next PO number


    With rsAgentTransferLines
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "AgentTransferLines", .ActiveConnection, , , adCmdTable
    End With

    With rstemp
        If gcfHandleErrors Then On Error GoTo ErrLogShowRstemp
        dtSetup = Now
        Do While Not .EOF
            nLineCount = nLineCount + 1
            rsAgentTransferLines.AddNew

            rsAgentTransferLines!AgentTransferHeadID = TransferID
            rsAgentTransferLines!LineID = nLineCount
            rsAgentTransferLines!ItemID = !ItemID
            rsAgentTransferLines!QuantityShipped = !SumOfQuantity
            If InOrOut = "I" Then
                rsAgentTransferLines!QuantityOrdered = !SumOfQuantity
                '!Status = "I"
            Else
                rsAgentTransferLines!QuantityOrdered = 0 - !SumOfQuantity
                '!Status = "O"
            End If

            rsAgentTransferLines!Price = !Price
            rsAgentTransferLines!Cost = !Cost

            rsAgentTransferLines!SetupDate = dtSetup
            rsAgentTransferLines!ChangeDate = dtSetup
            rsAgentTransferLines.Update
            If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SumOfQuantity
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
    Set comQDF = Nothing
    'delete the work file
    SetOptions False
    DoCmd.OpenQuery "qryMetrologicScanBatchDelete"
    SetOptions True
    CreateStoreTransferFromMetrologic = TransferID

ExitSub:
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!ItemID
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function




'Public Function PostPO_ADO(PurchaseOrderID As Long, VendorName As String)
'
'
'
'Dim strMsg As String
'Dim strSQL  As String
'
'
'Dim nLineCount As Integer
'Dim nLastInvoiceNo As Long
'Dim nShiftCounter  As Integer
'
'Dim rsPurchaseOrderLine As New ADODB.Recordset
'Dim rsInvoLine As New ADODB.Recordset
'Dim RSItem As New ADODB.Recordset
'Dim rsCashsys As New ADODB.Recordset
'
'Dim comQDF As New ADODB.Command
'
'
'
'    If gcfHandleErrors Then On Error GoTo ErrLog
'
'    strSQL = "SELECT * "
'    strSQL = strSQL & "FROM PurchaseOrderLine "
'    strSQL = strSQL & "WHERE ((PurchaseOrderID) = (" & PurchaseOrderID & "))"
'    rsPurchaseOrderLine.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'
'    If rsPurchaseOrderLine.RecordCount = 0 Then
'        rsPurchaseOrderLine.Close
'        Exit Function
'    End If
'
'    strSQL = "SELECT * "
'    strSQL = strSQL & "FROM CashSys "
'    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
'    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'
'    If rsCashsys.BOF And rsCashsys.EOF Then
'        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
'        Exit Function
'    End If
'
'    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
'    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 2000000
'    nShiftCounter = rsCashsys!ShiftCounter
'    rsCashsys.Update
'    rsCashsys.Close
'    Set rsCashsys = Nothing
'
'    strSQL = "Delete from PurchaseOrder "
'    strSQL = strSQL & "Where ((ID) = (" & CStr(PurchaseOrderID) & "))"
'
'    With comQDF
'        .ActiveConnection = CurrentProject.Connection
'        .CommandText = "qryPurchaseOrderPost_Step1_Update_Onhands"
'        .CommandType = adCmdStoredProc
'        .Parameters.Append .CreateParameter("parPurchaseOrderID", adInteger, adParamInput)
'        .Parameters("parPurchaseOrderID") = PurchaseOrderID
'        .Execute
'    End With
'    Set comQDF = Nothing
'
'    With comQDF
'        .ActiveConnection = CurrentProject.Connection
'        .CommandText = "qryPurchaseOrderPost_Step1_Update_Onhands"
'        .CommandType = adCmdStoredProc
'
'        .Parameters.Append .CreateParameter("parInvoiceNumber", adInteger, adParamInput)
'        .Parameters("parInvoiceNumber") = nLastInvoiceNo
'
'        .Parameters.Append .CreateParameter("parCashSysID", adInteger, adParamInput)
'        .Parameters("parCashSysID") = glnCashsysID
'
'        .Parameters.Append .CreateParameter("parCashierID", adInteger, adParamInput)
'        .Parameters("parCashierID") = glnCashierID
'
'        .Parameters.Append .CreateParameter("parShiftID", adInteger, adParamInput)
'        .Parameters("parShiftID") = nShiftCounter
'
'        .Parameters.Append .CreateParameter("parPurchaseOrderID", adInteger, adParamInput)
'        .Parameters("parPurchaseOrderID") = PurchaseOrderID
'        .Execute
'    End With
'    Set comQDF = Nothing
'
'
'
'
'    rsInvoLine.Open "InvoLine", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
'    RSItem.Open "Item", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
'
'    With rsPurchaseOrderLine
'        .MoveLast
'        SysCmd acSysCmdInitMeter, "Posting Lines", .RecordCount
'
'        .MoveFirst
'
'        Do While Not .EOF
'            SysCmd acSysCmdUpdateMeter, .AbsolutePosition
'            If !QuantityShipped <> 0 Then
'                nLineCount = nLineCount + 1
'                If nLineCount = 1 Then 'save Header
'                End If
'                RSItem.MoveFirst
'                RSItem.Find "((ID)=(" & CStr(!ItemID) & "))"
'                If RSItem.EOF Then
'                    MsgBox "Item ID" & CStr(!ItemID) & " not found!, Can not continue posting", vbCritical
'                    Exit Function
'                End If
'                RSItem!Onhand = RSItem!Onhand + !QuantityShipped
'                If !Cost <> 0 Then RSItem!Cost = !Cost
'                'added 09/05/2004
'                RSItem!ChangeDate = Now
'                RSItem.Update
'
'                rsInvoLine.AddNew
'                rsInvoLine!InvoHeadID = nLastInvoiceNo
'                rsInvoLine!ID = nLineCount
'                rsInvoLine!CashSysID = glnCashsysID
'                rsInvoLine!CashierID = glnCashierID
'                rsInvoLine!ShiftID = nShiftCounter
'                rsInvoLine!Code = !ItemID
'                rsInvoLine!SalesDept = RSItem!SalesDepartmentID
'                rsInvoLine!Quantity = !QuantityShipped
'                'changed 12/28/2004
'                rsInvoLine!Price = !QuantityShipped * !Price
'                rsInvoLine!InputCode = Left(CStr(!PurchaseOrderID) & " " & VendorName, 15)
'                'drawer
'                rsInvoLine!SalesGroup = RSItem!Onhand
'                rsInvoLine!TransType = "P" 'Purchase
'                'UpdateOnhands
'                rsInvoLine!SetupDate = Now
'                rsInvoLine.Update
'
'                '.Edit
'                !QuantityShipped = 0
'                .Update
'            End If
'            'If not gcfHandleErrors then Debug.Print !PurchaseOrderID, !VendorID, !LineID, !ItemID, !QuantityShipped, !Price, !Cost
'            .MoveNext
'        Loop
'        .Close
'    End With
'    rsInvoLine.Close
'    RSItem.Close
'
'    strSQL = "Delete from PurchaseOrder "
'    strSQL = strSQL & "Where ((ID) = (" & CStr(PurchaseOrderID) & "))"
'    With comQDF
'        .ActiveConnection = CurrentProject.Connection
'        .CommandType = adCmdText
'        .CommandText = strSQL
'        .Execute
'    End With
'
'ExitSub:
'    Set rsInvoLine = Nothing
'    Set RSItem = Nothing
'    Set rsPurchaseOrderLine = Nothing
'    Set comQDF = Nothing
'    SysCmd acSysCmdClearStatus
'    'UpdateMovement (True)
'    SQLPassThru "sp_UpdateItemMovement"
'
'
'    Exit Function
'
'ErrLog:
'    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Function:PostPO"
'    Resume ExitSub
'End Function

Public Function PostPO(PurchaseOrderID As Long, VendorName As String)

Dim strMsg As String
Dim strSQL  As String


'Dim rsCashsys As New ADODB.Recordset
Dim rsCashsys As Object
Set rsCashsys = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


'Dim adoCnOLCNetwork As New ADODB.Connection
Dim adoCnOLCNetwork As Object
Set adoCnOLCNetwork = CreateObject("ADODB.Connection")


Dim nLastInvoiceNo As Long
Dim nShiftCounter  As Integer
Dim qdfTemp As QueryDef

    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not gcfHandleErrors Then Debug.Print Now
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsCashsys.BOF And rsCashsys.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
        Exit Function
    End If

    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 2000000
    nShiftCounter = rsCashsys!ShiftCounter
    rsCashsys.Update
    rsCashsys.Close
    Set rsCashsys = Nothing

   '4/11/2014
    With adoCnOLCNetwork
        .ConnectionTimeout = 10
        .ConnectionString = "Provider=SQLNCLI.1;" _
        & "SERVER=" & gstrCashsysInsLoc & ";" _
        & "Database=KMDSSQL7;" _
        & "DataTypeCompatibility=80;" _
        & "Trusted_Connection=Yes;"
        .Open
    End With


    strSQL = "sp_PostPO "
    strSQL = strSQL & CStr(glnCashierID) & ", "
    strSQL = strSQL & CStr(glnCashsysID) & ", "
    strSQL = strSQL & CStr(nLastInvoiceNo) & ", "
    strSQL = strSQL & CStr(nShiftCounter) & ", "
    strSQL = strSQL & CStr(PurchaseOrderID)
    With comQDF
        .CommandTimeout = 15
        .ActiveConnection = adoCnOLCNetwork
        .CommandText = "sp_PostPO"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parCashierID", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("parCashsysID", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("parInvoHeadID", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("parShiftID", adInteger, adParamInput)
        .Parameters.Append .CreateParameter("parPurchaseOrderID", adInteger, adParamInput)
        .Parameters("parCashierID") = glnCashierID
        .Parameters("parCashsysID") = glnCashsysID
        .Parameters("parInvoHeadID") = nLastInvoiceNo
        .Parameters("parShiftID") = nShiftCounter
        .Parameters("parPurchaseOrderID") = PurchaseOrderID
        .Execute
    End With
    Set comQDF = Nothing
    adoCnOLCNetwork.Close
    Set adoCnOLCNetwork = Nothing

    If Not gcfHandleErrors Then Debug.Print Now
    SQLPassThru "sp_UpdateItemMovement"

    'UpdateMovement (True)
    If Not gcfHandleErrors Then Debug.Print Now
    AppendUPCItemsWithItemID

ExitSub:

    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Function:PostPO"
    Resume ExitSub
End Function


Public Function CreateTransfer(AgentID As Long, InOrOut As String) As Integer


'Dim rsAgent As New ADODB.Recordset
Dim rsAgent As Object
Set rsAgent = CreateObject("ADODB.Recordset")

'Dim rsAgentTransfer As New ADODB.Recordset
Dim rsAgentTransfer As Object
Set rsAgentTransfer = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim nOrderCount As Integer
Dim nTransferID As Integer
Dim nLineCount As Integer
Dim dtSetup As Date

If gcfHandleErrors Then On Error GoTo ErrLog
    CreateTransfer = 0

    'get the next AgentTransfer number
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryAgentTransferCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parAgentID", adInteger, adParamInput)
        .Parameters("parAgentID") = AgentID
    End With

    With rsAgent
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        !OrderCount = !OrderCount + 1
        nOrderCount = !OrderCount
        .Update
        .Close
        Set rsAgent = Nothing
    End With

    With rsAgentTransfer
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "AgentTransferHead", .ActiveConnection, , , adCmdTable
        .AddNew
        !AgentID = AgentID
        !Order = nOrderCount
        !OrdDate = Now
        !Status = InOrOut
        .Update
        .MoveLast
        nTransferID = !ID
        .Close
       Set rsAgentTransfer = Nothing
    End With

ExitSub:
    CreateTransfer = nTransferID
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function

Public Function PostTransfer(TransferID As Long, LastName As String) ', InOROut As String)

Dim strMsg As String
Dim strSQL  As String
Dim strInOrOut As String


Dim nLineCount As Integer
Dim nLastInvoiceNo As Long
Dim nShiftCounter  As Integer

'Dim rsPurchaseOrderHead As New ADODB.Recordset
Dim rsPurchaseOrderHead As Object
Set rsPurchaseOrderHead = CreateObject("ADODB.Recordset")

'Dim rsPurchaseOrderLine As New ADODB.Recordset
Dim rsPurchaseOrderLine As Object
Set rsPurchaseOrderLine = CreateObject("ADODB.Recordset")

'Dim rsInvoLine As New ADODB.Recordset
Dim rsInvoLine As Object
Set rsInvoLine = CreateObject("ADODB.Recordset")

'Dim RSItem As New ADODB.Recordset
Dim RSItem As Object
Set RSItem = CreateObject("ADODB.Recordset")

'Dim rsCashsys As New ADODB.Recordset
Dim rsCashsys As Object
Set rsCashsys = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nQuantity As Long
Dim nSaveAgentID As Long
Dim datNow As Date


    strMsg = "Are you ready to post this Transfer to Inventory?" & vbCrLf & vbCrLf
    strMsg = strMsg & "Yes = Add this order to onhands and move order to history file." & vbCrLf
    strMsg = strMsg & "No = Do Nothing; leave this order as is"
    If MsgBox(strMsg, vbYesNo, "Post Transfer #" & TransferID) = vbNo Then
        Exit Function
    End If
    If gcfHandleErrors Then On Error GoTo ErrLog


    strSQL = "SELECT * "
    strSQL = strSQL & "FROM AgentTransferHead "
    strSQL = strSQL & "WHERE ((ID) = (" & TransferID & "))"

    rsPurchaseOrderHead.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rsPurchaseOrderHead.RecordCount = 0 Then
        MsgBox "Not found Transfer #" & TransferID
        rsPurchaseOrderHead.Close
        Exit Function
    End If
    rsPurchaseOrderHead.MoveFirst
    strInOrOut = rsPurchaseOrderHead!Status
    nSaveAgentID = rsPurchaseOrderHead!AgentID
    rsPurchaseOrderHead.Close

    strSQL = "SELECT * "
    strSQL = strSQL & "FROM AgentTransferLines "
    strSQL = strSQL & "WHERE ((AgentTransferHeadID) = (" & TransferID & "))"
    rsPurchaseOrderLine.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsPurchaseOrderLine.RecordCount = 0 Then
        MsgBox "No lines found for Transfer #" & TransferID
        rsPurchaseOrderLine.Close
        Exit Function
    End If

    strSQL = "SELECT * "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsCashsys.BOF And rsCashsys.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
        Exit Function
    End If

    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 3000000
    nShiftCounter = rsCashsys!ShiftCounter
    datNow = Now
    rsCashsys.Update
    rsCashsys.Close
    Set rsCashsys = Nothing

    rsInvoLine.Open "InvoLine", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    RSItem.Open "Item", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable

    With rsPurchaseOrderLine
        .MoveLast
        SysCmd acSysCmdInitMeter, "Posting Lines", .RecordCount

        .MoveFirst

        Do While Not .EOF
            SysCmd acSysCmdUpdateMeter, .AbsolutePosition
            If !QuantityShipped <> 0 Then
                nLineCount = nLineCount + 1
                If nLineCount = 1 Then 'save Header
                End If
                If strInOrOut = "I" Then
                    nQuantity = !QuantityShipped
                Else
                    nQuantity = 0 - !QuantityShipped
                End If

                RSItem.MoveFirst
                RSItem.Find "((ID)=(" & CStr(!ItemID) & "))"
                If RSItem.EOF Then
                    MsgBox "Item ID" & CStr(!ItemID) & " not found!, Can not continue posting", vbCritical
                    Exit Function
                End If
                RSItem!Onhand = RSItem!Onhand + nQuantity '!QuantityShipped
                'changed 08/31/2004
                'rsItem!TransferredToday = rsItem!TransferredToday + nQuantity '!QuantityShipped
                'changed 09/05/2004
                RSItem!ChangeDate = datNow

                RSItem.Update

                rsInvoLine.AddNew
                rsInvoLine!InvoHeadID = nLastInvoiceNo
                rsInvoLine!ID = nLineCount
                rsInvoLine!CashSysID = glnCashsysID
                rsInvoLine!CashierID = glnCashierID
                rsInvoLine!ShiftID = nShiftCounter
                rsInvoLine!Code = !ItemID
                '09-24-2006
                rsInvoLine!SalesDept = nSaveAgentID 'RSItem!SalesDepartmentID

                rsInvoLine!Quantity = nQuantity '!QuantityShipped
                'changed 12/28/2004
                rsInvoLine!Price = nQuantity * !Price
                rsInvoLine!Cost = nQuantity * RSItem!Cost
                Select Case strInOrOut
                Case "C" ' claim
                    rsInvoLine!InputCode = !ClaimCode
                    'use as flag for claim for reprints!
                    rsInvoLine!Drawer = 1
                Case Else
                    rsInvoLine!InputCode = Left(CStr(nSaveAgentID) & " " & LastName & " (" & strInOrOut & ")", 15)    '!Order
                    rsInvoLine!Drawer = 0
                End Select
                'drawer
                'ADDED 2007-12-29
                rsInvoLine!SalesGroup = TransferID  ' RSItem!Onhand
                rsInvoLine!TransType = "X" 'Transfer
                'UpdateOnhands
                '2006-09-24 changed from Now to datNow so everyline Involine has same setupdate
                rsInvoLine!SetupDate = datNow
                '2006-09-24 drawer = 0 indicates a store transfer
                '               used by the 44-Transfers report selection!..
                rsInvoLine.Update

                '.Edit
                !QuantityShipped = 0
                .Update
            End If
            'If not gcfHandleErrors then Debug.Print !TransferID, !VendorID, !LineID, !ItemID, !QuantityShipped, !Price, !Cost
            .MoveNext
        Loop
        .Close
    End With
    rsInvoLine.Close
    RSItem.Close

    strSQL = "Delete from AgentTransferHead "
    strSQL = strSQL & "Where ((ID) = (" & CStr(TransferID) & "))"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Execute
    End With

ExitSub:
    Set rsInvoLine = Nothing
    Set RSItem = Nothing
    Set rsPurchaseOrderLine = Nothing
    Set comQDF = Nothing
    SysCmd acSysCmdClearStatus
    SQLPassThru "sp_UpdateItemMovement"
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Function:PostTransfer"
    Resume ExitSub
End Function


Public Function AssignDaysSupplyGroupIDs20051101()

'1. qryDaysSupplyStep1_Reset_GroupLevel_ADS             - set ADS and ADSGroup to 0
'2. qryDaysSupplyStep2_SetAverageDailySales             - set ADS
'3. qryDaysSupplyStep3_MakeDaysSupplyADSForNewItems     - maketable ADS for New Items
'4. qryDaysSupplyStep4_SetAverageDailySalesForNewItems  - set ADS for New Items
'5. qryDaysSupplyStep5_MakeDaysSupplyItemSales          - Make table of DaysSupplyItemSales
'7. qryDaysSupplyStep6_SetGroupLevel                    - Loop through DaysSupplyGroupLevels to assign group levels




'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim varReturn As Variant

    If gcfHandleErrors Then On Error Resume Next
    varReturn = SysCmd(acSysCmdSetStatus, "Deleting DaysSupplyItemMovementADSForNewItems")
    DoCmd.DeleteObject acTable, "DaysSupplyADSForNewItems"
    varReturn = SysCmd(acSysCmdSetStatus, "Deleting DaysSupplyItemSales")
    DoCmd.DeleteObject acTable, "DaysSupplyItemSales"

    If gcfHandleErrors Then On Error GoTo ErrLog

    varReturn = SysCmd(acSysCmdSetStatus, "Step #1 of 5. Execute qryDaysSupplyStep1_Reset_GroupLevel_ADS")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryDaysSupplyStep1_Reset_GroupLevel_ADS"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    varReturn = SysCmd(acSysCmdSetStatus, "Step #2 of 5. Execute qryDaysSupplyStep2_SetAverageDailySales")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryDaysSupplyStep2_SetAverageDailySales"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    varReturn = SysCmd(acSysCmdSetStatus, "Step #3 of 5. Execute qryDaysSupplyStep3_MakeDaysSupplyADSForNewItems")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryDaysSupplyStep3_MakeDaysSupplyADSForNewItems"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If gcfHandleErrors Then On Error GoTo ErrLog
    varReturn = SysCmd(acSysCmdSetStatus, "Step #4 of 5. Execute qryDaysSupplyStep4_SetAverageDailySalesForNewItems")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryDaysSupplyStep4_SetAverageDailySalesForNewItems"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    varReturn = SysCmd(acSysCmdSetStatus, "Step #5 of 5. qryDaysSupplyStep5_MakeDaysSupplyItemSales")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryDaysSupplyStep5_MakeDaysSupplyItemSales"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    With rstemp
        .Open "DaysSupplyGroupLevels", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
        Do While Not .EOF
            If Not gcfHandleErrors Then Debug.Print !GroupLevel; !MinUnitSales; !MaxUnitSales
            varReturn = SysCmd(acSysCmdSetStatus, "qryDaysSupplyStep6_SetGroupLevel Assigning group Levels " & !GroupLevel & " " & !MinUnitSales & " " & !MaxUnitSales)
            With comQDF
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryDaysSupplyStep6_SetGroupLevel"
                .CommandType = adCmdStoredProc
                'ADO Numeric Datatypes are very particular
                .Parameters.Append .CreateParameter("parDaysSupplyGroupLevel", adBigInt, adParamInput)
                .Parameters("parDaysSupplyGroupLevel") = rstemp!GroupLevel
                .Parameters.Append .CreateParameter("parLowerAmount", adBigInt, adParamInput)
                .Parameters("parLowerAmount") = rstemp!MinUnitSales
                .Parameters.Append .CreateParameter("parUpperAmount", adBigInt, adParamInput)
                .Parameters("parUpperAmount") = rstemp!MaxUnitSales
                .Execute
            End With
            Set comQDF = Nothing
            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing

    varReturn = SysCmd(acSysCmdClearStatus)

ExitSub:
    Set comQDF = Nothing
    SysCmd acSysCmdClearStatus
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Function:AssignDaysSupplyGroupIDs"
    Resume ExitSub
End Function


Public Function AssignDaysSupplyGroupIDs()

'1. qryADS_Step_1_Reset_DaysSupplyGroupID             - set ADS and ADSGroup to 0
'2. qryADS_Step_2_Set_AverageDailySales             - set ADS
'3. qryADS_Step_3_Make_DaysSupplyADSForNewItems     - maketable ADS for New Items
'4. qryADS_Step_4_Set_ADSForNewItems  - set ADS for New Items
'5. qryADS_Step_5_Make_DaysSupplyItemSales          - Make table of DaysSupplyItemSales
'6. qryADS_Step_6_AssignDaysSupplyGroupID                    - Loop through DaysSupplyGroupLevels to assign group levels


'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")



'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim comQDF2 As New Command
Dim comQDF2 As Object
Set comQDF2 = CreateObject("ADODB.COMMAND")


Dim varReturn As Variant

    If gcfHandleErrors Then On Error Resume Next
    varReturn = SysCmd(acSysCmdSetStatus, "Deleting DaysSupplyItemMovementADSForNewItems")
    DoCmd.DeleteObject acTable, "DaysSupplyADSForNewItems"
    varReturn = SysCmd(acSysCmdSetStatus, "Deleting DaysSupplyItemSales")
    DoCmd.DeleteObject acTable, "DaysSupplyItemSales"

    If gcfHandleErrors Then On Error GoTo ErrLog

    varReturn = SysCmd(acSysCmdSetStatus, "Step #1 of 5. Execute qryADS_Step_1_Reset_DaysSupplyGroupID")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryADS_Step_1_Reset_DaysSupplyGroupID"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    varReturn = SysCmd(acSysCmdSetStatus, "Step #2 of 5. Execute qryADS_Step_2_Set_AverageDailySales")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryADS_Step_2_Set_AverageDailySales"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    varReturn = SysCmd(acSysCmdSetStatus, "Step #3 of 5. Execute qryADS_Step_3_Make_DaysSupplyADSForNewItems")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryADS_Step_3_Make_DaysSupplyADSForNewItems"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    If gcfHandleErrors Then On Error GoTo ErrLog
    varReturn = SysCmd(acSysCmdSetStatus, "Step #4 of 5. Execute qryADS_Step_4_Set_ADSForNewItems")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryADS_Step_4_Set_ADSForNewItems"
        .CommandType = adCmdStoredProc
'        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    varReturn = SysCmd(acSysCmdSetStatus, "Step #5 of 5. qryADS_Step_5_Make_DaysSupplyItemSales")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryADS_Step_5_Make_DaysSupplyItemSales"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryADS_VendorsWithAverageDaySupplyID"
        .CommandType = adCmdStoredProc
    End With

    Debug.Print Now
    With rstemp
        .Open comQDF, , adOpenStatic, adLockReadOnly
        Do While Not .EOF
            If Not gcfHandleErrors Then Debug.Print !VendorID; !AverageDaySupplyID; !MinUnitSales; !MaxUnitSales
            varReturn = SysCmd(acSysCmdSetStatus, "Step_6 AssignDaysSupplyGroupID " & !VendorID & " " & !AverageDaySupplyID & " " & !MinUnitSales & " " & !MaxUnitSales)
            With comQDF2
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryADS_Step_6_AssignDaysSupplyGroupID"
                .CommandType = adCmdStoredProc
                'ADO Numeric Datatypes are very particular
                .Parameters.Append .CreateParameter("parDaysSupplyGroupLevel", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parLowerAmount", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parUpperAmount", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
                .Parameters("parDaysSupplyGroupLevel") = rstemp!AverageDaySupplyLevelsID
                .Parameters("parLowerAmount") = rstemp!MinUnitSales
                .Parameters("parUpperAmount") = rstemp!MaxUnitSales
                .Parameters("parVendorID") = rstemp!VendorID
                .Execute
            End With
            Set comQDF2 = Nothing
            Set comQDF2 = CreateObject("ADODB.COMMAND")
            .MoveNext
        Loop
        .Close
    End With
    Debug.Print Now
    Set rstemp = Nothing
    Set comQDF = Nothing

    varReturn = SysCmd(acSysCmdClearStatus)

ExitSub:
    Set comQDF = Nothing
    SysCmd acSysCmdClearStatus
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Function:AssignDaysSupplyGroupIDs"
    Resume ExitSub
End Function




Public Sub OLCCElectronicPackingSlipImport()

On Error GoTo ErrLog

Dim strTemp As String
Dim strDestError As String
Dim strQList As String
Dim strEPSOrderDate  As String
Dim strEPSAgent As String
Dim strEPSOrder As String
Dim strSourceDir As String
Dim strEPSOrderYYYY As String
Dim strEPSOrderMM As String
Dim strEPSOrderDD As String
Dim strSourceFname
Dim strDestFName
Dim strSQL As String
Dim strDateStamp As String


Dim fsoTemp As New FileSystemObject


    strTemp = "C:\NPOS\Mail"

    'lets build a structure ..
    If Not fsoTemp.FolderExists(strTemp) Then
        fsoTemp.CreateFolder (strTemp)
    End If

    If Not fsoTemp.FolderExists(strTemp & "\In") Then
        fsoTemp.CreateFolder (strTemp & "\In")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\In\Archive") Then
        fsoTemp.CreateFolder (strTemp & "\In\Archive")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\In\ERROR") Then
        fsoTemp.CreateFolder (strTemp & "\In\ERROR")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\In\Temp") Then
        fsoTemp.CreateFolder (strTemp & "\In\Temp")
    End If


    SetOptions False
    strSourceDir = "C:\NPOS\MAIL\IN\"

    strTemp = Dir(strSourceDir & "PackingList." & CStr(glnAgentNumber) & ".*.txt")
    If strTemp <> "" Then
        DoCmd.RunSQL "Delete * from OLCCEPS", False
    End If

    Do While strTemp <> vbNullString
        strEPSOrderMM = Mid(strTemp, 34, 2)
        strEPSOrderDD = Mid(strTemp, 37, 2)
        strEPSOrderYYYY = Mid(strTemp, 29, 4)

        strEPSOrderDate = strEPSOrderMM & "/" & strEPSOrderDD & "/" & strEPSOrderYYYY
        strEPSAgent = Mid(strTemp, 13, 4)
        strEPSOrder = Mid(strTemp, 18, 10)

        strSourceFname = strSourceDir & strTemp
        strDestFName = strSourceDir & "EPS" & strEPSAgent & strEPSOrder & strEPSOrderMM & strEPSOrderDD & strEPSOrderYYYY & ".TXT"

        FileCopy strSourceFname, strDestFName
        Kill strSourceFname
        DoCmd.TransferText acImportFixed, "EPSImportSpecification", "OLCCEPS", strDestFName, False
        'add the order date
        strSQL = "Update OLCCEPS Set OrderDate = #" & strEPSOrderDate & "# "
        strSQL = strSQL & "WHERE OrderNumber = " & strEPSOrder & " AND OrderDate = 0"
        DoCmd.RunSQL strSQL

        strDateStamp = Format(Now, "yyyy-mm-dd-hh-nn")

        strSourceFname = strDestFName
        strDestFName = strSourceDir & "Archive\EPS" & strEPSAgent & strEPSOrder & strEPSOrderMM & strEPSOrderDD & strEPSOrderYYYY & " " & strDateStamp & ".TXT"
        FileCopy strSourceFname, strDestFName
        Kill strSourceFname


        strQList = strQList & strEPSOrderDate & " " & strEPSOrder & " " & strEPSAgent & vbCrLf
        'txtFileName = ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\" & strTemp)
        'ImportClubOrderFromExcel txtFileName
        'If fsotemp.FileExists(txtFileName) Then  'IT DIDN'T GET MOVED... MOVE IT TO ERROR FILE
        '    strDestError = Replace(txtFileName, "\IN\", "\IN\ERROR\" & Format(Now, "yyyymmdd hh nn ss"))
        '    fsotemp.MoveFile txtFileName, strDestError
        'End If
        strTemp = Dir
    Loop



ExitSub:
    SetOptions True
Exit Sub



ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Function:OLCCElectronicPackingSlipImport"
    Resume ExitSub

End Sub
