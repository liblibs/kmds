Attribute VB_Name = "modClubOrders"
Option Compare Database
Option Explicit
Dim mnLinecount As Long

Public Function CreateClubOrder(ClubID As Long, OrderGuide As Boolean, ClubIDForOrderGuide As Long)

'Dim rstemp As New ADODB.Recordset
Dim rsClub As Object
Set rsClub = CreateObject("ADODB.Recordset")

'Dim rstemp As New ADODB.Recordset
Dim rsOrder As Object
Set rsOrder = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


Dim nOrderCount As Integer
Dim nOrderID As Long
Dim bNonLiquor As Boolean
Dim strSaveNotes As String

If gcfHandleErrors Then On Error GoTo ErrLog
    CreateClubOrder = 0
    'ArchiveReOrderrt "OrderCreate"


    'get the next Order number
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
        .Parameters("parClubID") = ClubID
    End With

    With rsClub
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        .MoveFirst
        nOrderCount = !OrderCount
        !OrderCount = !OrderCount + 1
        bNonLiquor = !NonLiquor
        If IsNull(!ClubNotes) Then
        Else
            strSaveNotes = !ClubNotes
        End If
        .Update
        .Close
    End With


    Set rsClub = Nothing

    With rsOrder
        Set .ActiveConnection = CurrentProject.Connection
        .CursorLocation = adUseServer
        .CursorType = adOpenKeyset
        .LockType = adLockPessimistic
        .Index = "PrimaryKey"
        .Open "OrdHead", .ActiveConnection, adOpenKeyset, adLockOptimistic
        .AddNew
        !ClubID = ClubID
        !Order = nOrderCount
        '188    2006-09-22  Order Notes default to the customer notes in CreateClubOrder routine

        !OrderNotes = strSaveNotes
        .Update
        .MoveLast
        nOrderID = !ID
        .Close
       Set rsOrder = Nothing
    End With

    AssignClubInvoiceNO CLng(nOrderID)

    If OrderGuide = True Then
    Dim bOrderGuideSeparateRelated As Boolean
        bOrderGuideSeparateRelated = GetKMDSSettings("OrderGuideSeparateRelated", -1)
        If bOrderGuideSeparateRelated Then
            mnLinecount = 0
            CreateClubOrderLines ClubIDForOrderGuide, nOrderID, True
            CreateClubOrderLines ClubIDForOrderGuide, nOrderID, False
        Else
            CreateClubOrderLinesLiquorRelated ClubIDForOrderGuide, nOrderID
        End If
    Else


    End If


ExitSub:
    CreateClubOrder = nOrderID
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    mnLinecount = -1
    Resume ExitSub
End Function


Public Sub CreateClubOrderLines(ClubIDForOrderGuide As Long, OrderID As Long, LiquorItems As Boolean)
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rsOrderLine As New ADODB.Recordset
Dim rsOrderLine As Object
Set rsOrderLine = CreateObject("ADODB.Recordset")

Dim dtSetup As Date

        With rsOrderLine
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Index = "PrimaryKey"
            .Open "OrderLine", .ActiveConnection, , , adCmdTable
        End With

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryClubOrderGuides"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
            .Parameters("parClubID") = ClubIDForOrderGuide
            .Parameters.Append .CreateParameter("parLiquor", adBoolean, adParamInput)
            .Parameters("parLiquor") = LiquorItems
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
        With rstemp
            dtSetup = Now
            If Not gcfHandleErrors Then Debug.Print Now
            Do While Not .EOF
                mnLinecount = mnLinecount + 1
                rsOrderLine.AddNew
                'rsOrderLine!ClubID = ClubID
                rsOrderLine!OrdheadID = OrderID
                rsOrderLine!LineID = mnLinecount
                rsOrderLine!ItemID = !ItemID
                'rsOrderLine!QuantityShipped 'Ordered = IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!QuantityOrdered = !Par 'IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!Price = !Sprice '+ !Buydown
                rsOrderLine!SalePrice = IIf(!SP, "*", "")
                If Not LiquorItems And !QuotedPrice > 0 Then
                    rsOrderLine!Price = !QuotedPrice
                    rsOrderLine!SalePrice = "Q"
                End If
                rsOrderLine!Cost = !Cost
                rsOrderLine!Buydown = !Buydown
                rsOrderLine!SetupDate = dtSetup
                rsOrderLine!ChangeDate = dtSetup
                rsOrderLine.Update
                .MoveNext
            Loop
            .Close
            If Not gcfHandleErrors Then Debug.Print Now
        End With
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set comQDF = Nothing
    Set rsOrderLine = Nothing

ExitSub:
    Exit Sub
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    'mnLinecount = -1
    Resume ExitSub
End Sub

Public Sub CreateClubOrderLinesLiquorRelated(ClubIDForOrderGuide As Long, OrderID As Long)
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rsOrderLine As New ADODB.Recordset
Dim rsOrderLine As Object
Set rsOrderLine = CreateObject("ADODB.Recordset")

Dim dtSetup As Date

        With rsOrderLine
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Index = "PrimaryKey"
            .Open "OrderLine", .ActiveConnection, , , adCmdTable
        End With

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryClubOrderGuidesLiquorRelated"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
            .Parameters("parClubID") = ClubIDForOrderGuide
            .Parameters.Append .CreateParameter("parLiquor", adBoolean, adParamInput)
            '.Parameters("parLiquor") = LiquorItems
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
        With rstemp
            dtSetup = Now
            If Not gcfHandleErrors Then Debug.Print Now
            Do While Not .EOF
                mnLinecount = mnLinecount + 1
                rsOrderLine.AddNew
                'rsOrderLine!ClubID = ClubID
                rsOrderLine!OrdheadID = OrderID
                rsOrderLine!LineID = mnLinecount
                rsOrderLine!ItemID = !ItemID
                'rsOrderLine!QuantityShipped 'Ordered = IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!QuantityOrdered = !Par 'IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!Price = !Sprice '+ !Buydown
                rsOrderLine!SalePrice = IIf(!SP, "*", "")
'                If Not LiquorItems And !QuotedPrice > 0 Then
'                    rsOrderLine!Price = !QuotedPrice
'                    rsOrderLine!SalePrice = "Q"
'                End If
                rsOrderLine!Cost = !Cost
                rsOrderLine!Buydown = !Buydown
                rsOrderLine!SetupDate = dtSetup
                rsOrderLine!ChangeDate = dtSetup
                rsOrderLine.Update
                .MoveNext
            Loop
            .Close
            If Not gcfHandleErrors Then Debug.Print Now
        End With
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set comQDF = Nothing
    Set rsOrderLine = Nothing

ExitSub:
    Exit Sub
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    'mnLinecount = -1
    Resume ExitSub
End Sub

Public Sub ClubOrderRecalcTotals(OrderID As Long)

'Dim RS As New ADODB.Recordset
Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = CurrentDb().OpenRecordset("select * from  qryClubOrderLineTotals where OrdheadID =" & OrderID, dbReadOnly)
    With rs
        If .RecordCount <> 0 Then
            .MoveFirst
            With Forms![ClubOrder]

                If IsNull(!txtInvDollars) Or !txtInvDollars <> rs!ExtPrice Then
                    !txtInvDollars = rs!ExtPrice
                End If
                If IsNull(!txtDisDollars) Or !txtDisDollars <> (rs!LiquorDiscount + rs!NonLiquorDiscount) Then
                    !txtDisDollars = (rs!LiquorDiscount + rs!NonLiquorDiscount)
                End If
                If IsNull(!txtOrdDollars) Or !txtOrdDollars <> rs!ExtNetPrice Then
                    !txtOrdDollars = rs!ExtNetPrice
                End If


                If IsNull(!txtInvQty) Or !txtInvQty <> (rs!LiquorQuantity + rs!NonLiquorQuantity) Then
                    !txtInvQty = rs!LiquorQuantity + rs!NonLiquorQuantity
                End If



                If IsNull(!txtLiquorDollars) Or !txtLiquorDollars <> rs!LiquorExtPrice Then
                    !txtLiquorDollars = rs!LiquorExtPrice
                End If

                If IsNull(!txtLiquorDiscount) Or !txtLiquorDiscount <> rs!LiquorDiscount Then
                    !txtLiquorDiscount = rs!LiquorDiscount
                End If

                If IsNull(!txtNonLiquorDollars) Or !txtNonLiquorDollars <> rs!NonLiquorExtPrice Then
                    !txtNonLiquorDollars = rs!NonLiquorExtPrice
                End If

                If IsNull(!txtNonLiquorDiscount) Or !txtNonLiquorDiscount <> rs!NonLiquorDiscount Then
                    !txtNonLiquorDiscount = rs!NonLiquorDiscount
                End If
                    'added 6/5/2018


                If .Form.Dirty Then
                    .Form.Dirty = False
                End If
            End With
        .Close
        End If
    End With
    Set rs = Nothing
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Club Order Recalc Totals"
    Resume ExitSub
End Sub

Public Function CreateClubOrderFromPDA(ClubID As Long) As Long

Dim nOrderID As Long
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rsOrderLine As New ADODB.Recordset
Dim rsOrderLine As Object
Set rsOrderLine = CreateObject("ADODB.Recordset")

Dim dtSetup As Date

    nOrderID = CreateClubOrder(ClubID, False, ClubID)
    If nOrderID > 0 Then 'no error !
        CreateClubOrderFromPDA = nOrderID
        With rsOrderLine
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Index = "PrimaryKey"
            .Open "OrderLine", .ActiveConnection, , , adCmdTable
        End With

        'Set comQDF = Nothing
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPSCClubOrder"
            .CommandType = adCmdStoredProc
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
        With rstemp
            dtSetup = Now
            If Not gcfHandleErrors Then Debug.Print Now
            Do While Not .EOF
                mnLinecount = mnLinecount + 1
                rsOrderLine.AddNew
                'rsOrderLine!ClubID = ClubID
                rsOrderLine!OrdheadID = nOrderID
                rsOrderLine!LineID = mnLinecount
                rsOrderLine!ItemID = !ID
                If !QuantityShipped = -1 Then
                    rsOrderLine!QuantityShipped = 1
                Else
                    rsOrderLine!QuantityShipped = !QuantityShipped
                End If
                'rsOrderLine!QuantityOrdered = !Par 'IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!Price = !CurPrice '+ !Buydown
'                rsOrderLine!SalePrice = IIf(!SP, "*", "")
'                If Not LiquorItems And !QuotedPrice > 0 Then
'                    rsOrderLine!Price = !QuotedPrice
'                    rsOrderLine!SalePrice = "Q"
'                End If
                rsOrderLine!Cost = !Cost
                rsOrderLine!Buydown = !Buydown
                rsOrderLine!SetupDate = dtSetup
                rsOrderLine!ChangeDate = dtSetup
                rsOrderLine.Update
                .MoveNext
            Loop
            .Close
            If Not gcfHandleErrors Then Debug.Print Now
        End With
    End If
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set rsOrderLine = Nothing

    'delete the work file
    SetOptions False
    DoCmd.OpenQuery "qryPSCInput_Step1_DeleteRecords"
    SetOptions True

ExitSub:
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    'mnLinecount = -1
    Resume ExitSub
End Function


Public Sub RepriceClubOrder()

Dim varReturn
If gcfHandleErrors Then On Error GoTo ErrLog

    'added 02/01/2005
    'This will 'reprice' all orders with vendor #1
    'to fix the price for orders input in the previous month
    SQLPassThru ("sp_Upd_OrderLinesWithCurrentPrice")
    'MsgBox "pt"
    'cmdReprice.Visible = False

ExitSub:
   varReturn = SysCmd(acSysCmdClearStatus)
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub


End Sub

Public Function CreateClubOrderFromMetroLogic(ClubID As Long) As Long

Dim nOrderID As Long

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rsOrderLine As New ADODB.Recordset
Dim rsOrderLine As Object
Set rsOrderLine = CreateObject("ADODB.Recordset")

Dim dtSetup As Date

    nOrderID = CreateClubOrder(ClubID, False, ClubID)
    If nOrderID > 0 Then 'no error !
        CreateClubOrderFromMetroLogic = nOrderID
        With rsOrderLine
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Index = "PrimaryKey"
            .Open "OrderLine", .ActiveConnection, , , adCmdTable
        End With

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryMetrologicScanBatchForOrders"
            .CommandType = adCmdStoredProc
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
        With rstemp
            dtSetup = Now
            If Not gcfHandleErrors Then Debug.Print Now
            Do While Not .EOF
                mnLinecount = mnLinecount + 1
                rsOrderLine.AddNew
                'rsOrderLine!ClubID = ClubID
                rsOrderLine!OrdheadID = nOrderID
                rsOrderLine!LineID = mnLinecount
                rsOrderLine!ItemID = !ItemID
                If !SumOfQuantity < 1 Then
                    rsOrderLine!QuantityShipped = 1
                Else
                    rsOrderLine!QuantityShipped = !SumOfQuantity
                End If
                'rsOrderLine!QuantityOrdered = !Par 'IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!Price = !Price '+ !Buydown
'                rsOrderLine!SalePrice = IIf(!SP, "*", "")
'                If Not LiquorItems And !QuotedPrice > 0 Then
'                    rsOrderLine!Price = !QuotedPrice
'                    rsOrderLine!SalePrice = "Q"
'                End If
                rsOrderLine!Cost = !Cost
'                rsOrderLine!Buydown = !Buydown
                rsOrderLine!SetupDate = dtSetup
                rsOrderLine!ChangeDate = dtSetup
                rsOrderLine.Update
                .MoveNext
            Loop
            .Close
            If Not gcfHandleErrors Then Debug.Print Now
        End With
    End If
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set comQDF = Nothing
    Set rsOrderLine = Nothing

    'delete the work file
    SetOptions False
    DoCmd.OpenQuery "qryMetrologicScanBatchDelete"
    SetOptions True

ExitSub:
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    'mnLinecount = -1
    Resume ExitSub
End Function



Public Function CreateClubOrderFromScanBatch(ClubID As Long) As Long

Dim nOrderID As Long

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rsOrderLine As New ADODB.Recordset
Dim rsOrderLine As Object
Set rsOrderLine = CreateObject("ADODB.Recordset")

Dim dtSetup As Date

    nOrderID = CreateClubOrder(ClubID, False, ClubID)
    If nOrderID > 0 Then 'no error !
        CreateClubOrderFromScanBatch = nOrderID
        With rsOrderLine
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Index = "PrimaryKey"
            .Open "OrderLine", .ActiveConnection, , , adCmdTable
        End With

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryScanBatchForOrders"
            .CommandType = adCmdStoredProc
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
        With rstemp
            dtSetup = Now
            If Not gcfHandleErrors Then Debug.Print Now
            Do While Not .EOF
                mnLinecount = mnLinecount + 1
                rsOrderLine.AddNew
                'rsOrderLine!ClubID = ClubID
                rsOrderLine!OrdheadID = nOrderID
                rsOrderLine!LineID = mnLinecount
                rsOrderLine!ItemID = !ItemID
                If !SumOfCountQuantity < 1 Then
                    rsOrderLine!QuantityShipped = 1
                Else
                    rsOrderLine!QuantityShipped = !SumOfCountQuantity
                End If
                'rsOrderLine!QuantityOrdered = !Par 'IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!Price = !FirstOfPrice '+ !Buydown
'                rsOrderLine!SalePrice = IIf(!SP, "*", "")
'                If Not LiquorItems And !QuotedPrice > 0 Then
'                    rsOrderLine!Price = !QuotedPrice
'                    rsOrderLine!SalePrice = "Q"
'                End If
                rsOrderLine!Cost = !FirstOfCost
'                rsOrderLine!Buydown = !Buydown
                rsOrderLine!SetupDate = dtSetup
                rsOrderLine!ChangeDate = dtSetup
                rsOrderLine.Update
                .MoveNext
            Loop
            .Close
            If Not gcfHandleErrors Then Debug.Print Now
        End With
    End If
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set comQDF = Nothing
    Set rsOrderLine = Nothing

    'delete the work file
    SetOptions False
    DoCmd.OpenQuery "qryPSCInput_Step1_DeleteRecords"
    SetOptions True

ExitSub:
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    'mnLinecount = -1
    Resume ExitSub
End Function


Public Function CreateClubOrderFromSalesHistory(ClubID As Long) As Long

Dim nOrderID As Long


'Dim rsOrderLine As New ADODB.Recordset
Dim rsOrderLine As Object
Set rsOrderLine = CreateObject("ADODB.Recordset")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim dtSetup As Date

    nOrderID = CreateClubOrder(ClubID, False, ClubID)
    If nOrderID > 0 Then 'no error !
        CreateClubOrderFromSalesHistory = nOrderID
        With rsOrderLine
            Set .ActiveConnection = CurrentProject.Connection
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Index = "PrimaryKey"
            .Open "OrderLine", .ActiveConnection, , , adCmdTable
        End With

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryClubOrderCreateFromSalesHistory"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parClubID", adInteger, adParamInput)
            .Parameters("parClubID") = ClubID
        End With

        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
        With rstemp
            dtSetup = Now
            If Not gcfHandleErrors Then Debug.Print Now
            Do While Not .EOF
                mnLinecount = mnLinecount + 1
                rsOrderLine.AddNew
                'rsOrderLine!ClubID = ClubID
                rsOrderLine!OrdheadID = nOrderID
                rsOrderLine!LineID = mnLinecount
                rsOrderLine!ItemID = !Code
                rsOrderLine!QuantityShipped = 0
                rsOrderLine!QuantityShipped = 0
                'rsOrderLine!QuantityOrdered = !Par 'IIf(!SOQ > 0, !SOQ, 0)
                rsOrderLine!Price = !LastOfPrice '+ !Buydown
'                rsOrderLine!SalePrice = IIf(!SP, "*", "")
'                If Not LiquorItems And !QuotedPrice > 0 Then
'                    rsOrderLine!Price = !QuotedPrice
'                    rsOrderLine!SalePrice = "Q"
'                End If
                rsOrderLine!Cost = !LastOfCost
'                rsOrderLine!Buydown = !Buydown
                rsOrderLine!SetupDate = dtSetup
                rsOrderLine!ChangeDate = dtSetup
                rsOrderLine.Update
                .MoveNext
            Loop
            .Close
            If Not gcfHandleErrors Then Debug.Print Now
        End With
    End If
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set comQDF = Nothing
    Set rsOrderLine = Nothing

ExitSub:
    Exit Function
ErrLog:

    MsgBox Err.Number & " " & Err.Description
    'mnLinecount = -1
    Resume ExitSub
End Function

Public Sub FinTechExport()


    Dim strFintechURL As String
    Dim strFintechUserName As String
    Dim strFintechPassword As String
    Dim strFileName As String
    Dim strDirName As String
    Dim strFinTechLogon As String
    Dim strTemp As String

    Dim fsoTemp As New FileSystemObject
    Dim txtstrTemp As TextStream

    If gcfHandleErrors Then On Error GoTo ErrLog

    ' build directory structure
    If Not fsoTemp.FolderExists("k:\mail") Then
        fsoTemp.CreateFolder ("k:\mail")
    End If
    If Not fsoTemp.FolderExists("k:\mail\FinTech") Then
        fsoTemp.CreateFolder ("k:\mail\FinTech")
    End If
    If Not fsoTemp.FolderExists("k:\mail\FinTech\Out") Then
        fsoTemp.CreateFolder ("k:\mail\FinTech\Out")
    End If
    If Not fsoTemp.FolderExists("k:\mail\FinTech\Out\Archive") Then
        fsoTemp.CreateFolder ("k:\mail\FinTech\Out\Archive")
    End If
    If Not fsoTemp.FolderExists("k:\mail\FinTech\In") Then
        fsoTemp.CreateFolder ("k:\mail\FinTech\In")
    End If
    If Not fsoTemp.FolderExists("k:\mail\FinTech\In\Archive") Then
        fsoTemp.CreateFolder ("k:\mail\FinTech\In\Archive")
    End If

    'build file name with date/time stamp

    strDirName = "k:\mail\FinTech\Out"
    strFileName = strDirName & "\FT" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & " " & Hour(Now) & "-" & Minute(Now) & ".csv"


    'use transfertext to create the file
    DoCmd.TransferText acExportDelim, , "qryFintechExport", strFileName, True


    'get Fintech Credentials
    strFintechURL = GetKMDSSettings("Fintech.URL", "ftp.kmds.us")
    strFintechUserName = GetKMDSSettings("Fintech.UserName", "Agents")
    strFintechPassword = GetKMDSSettings("Fintech.Password", "Ag3nt$")

    'build logon string used by kmdsftp
    strFinTechLogon = strFintechURL & vbCrLf
    strFinTechLogon = strFinTechLogon & strFintechUserName & vbCrLf
    strFinTechLogon = strFinTechLogon & strFintechPassword & vbCrLf


    'add MPUT command as 4th line of string (note 1 at the end will move put file to archive directory
    strFinTechLogon = strFinTechLogon & "MPUT" & vbTab & strDirName & vbTab & "FT2*.csv" & vbTab & IIf(gbOLCCFTPAscii, 1, 0) & vbTab & 1

    'create script file used by KMDSFTP
    Set txtstrTemp = fsoTemp.OpenTextFile(strDirName & "\FinTechPut.txt", ForWriting, True, TristateFalse)
    txtstrTemp.WriteLine strFinTechLogon
    txtstrTemp.Close

    'chain out to ftp
    WaitForExecCmd "\\" & gstrCashsysInsLoc & "\kmds\kmdsftp.exe " & strDirName & "\FinTechPut.txt" ', vbNormalFocus


ExitSub:
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "OLCCDailySalesFile"
    Resume ExitSub
End Sub

Public Function ExportClubInvoiceToCSV() As String


Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
'Dim rstemp As Object
'Dim comQDF As Object

'Dim txtstrtemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String
Dim nClubID As Long
Dim strDirName As String
Dim strFileName   As String


    If gcfHandleErrors Then On Error GoTo ErrLog
    'must finish completely to get a valid file name
    ExportClubInvoiceToCSV = "NoOutput"

        ' build directory structure
    If Not fsoTemp.FolderExists("k:\mail") Then
        fsoTemp.CreateFolder ("k:\mail")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out") Then
        fsoTemp.CreateFolder ("k:\mail\Out")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\ClubCSV") Then
        fsoTemp.CreateFolder ("k:\mail\Out\ClubCSV")
    End If
    If Not fsoTemp.FolderExists("k:\mail\Out\ClubCSV\Archive") Then
        fsoTemp.CreateFolder ("k:\mail\Out\ClubCSV\Archive")
    End If

    'build file name with date/time stamp

    strDirName = "k:\mail\Out\ClubCSV\"
    With Forms![ClubOrderPrintDialog]
        strFileName = !txtClubName & " - " & !txtInvoiceNo & " - "
    End With
    strFileName = strFileName & Format(Now, "yyyy-dd-mm hh-nn") & ".csv"

    SetOptions False
    DoCmd.OpenQuery "qryClubInvoiceForExport_01_Delete_Records"
    DoCmd.OpenQuery "qryClubInvoiceForExport_02_Append_Lines"
    DoCmd.OpenQuery "qryClubInvoiceForExport_03_Append_Discount_Liquor"
    DoCmd.OpenQuery "qryClubInvoiceForExport_04_Append_Discount_NonLiquor"
    'DoCmd.TransferText acExportDelim, , "ClubInvoiceForCSVExport", strDirName & strFileName, True
    strDirName = "k:\mail\Out\ClubCSV\Archive\"
    DoCmd.TransferText acExportDelim, , "ClubInvoiceForCSVExport", strDirName & strFileName, True

    SetOptions True
    'success - return filename to attach
    ExportClubInvoiceToCSV = strDirName & strFileName

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
'    Set comQDF = Nothing
'    Set rstemp = Nothing
'    Set txtstrtemp = Nothing
    Set fsoTemp = Nothing
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportOrderGuideFileForWeb"
    Resume ExitSub
End Function
