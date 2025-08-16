Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbTPR As Boolean

Private Sub chkFTPFZActive_Click()
    SetKMDSSettings "FTPFZActive", chkFTPFZActive

End Sub

Private Sub cmbWADSDesc_Click()
    MsgBox ("Change this value in menu option C-2" & vbCrLf & "Make sure to click the 'Update Average Daily Sales' button after changing value!")

End Sub

Private Sub cmbWADSDesc_Enter()
cmbWADSDesc_Click
End Sub

Private Sub cmdCreatePO_Click()
    Dim stDocName As String
    Dim nOrderID As Integer
    Dim stLinkCriteria As String

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")

    Dim datBegin As Date
    Dim datEnd As Date
    Dim ndays As Integer
    Dim qdfTemp As QueryDef

''    If gcfHandleErrors Then On Error GoTo ErrLog

    FormatDateRanges
    lbRetr.Caption = "Building New PO..."
    lbRetr.Visible = True
    Me.Repaint

    Select Case optBasePO
    Case 1
        stDocName = "PurchaseOrderCreate"
        'DoCmd.OpenReport stDocName, acViewPreview
        ArchiveReport stDocName
        nOrderID = CreatePurchaseOrder(lstVendors, chkSaleItems.Value, chkTPR.Value)
        If nOrderID > 0 Then
            stDocName = "PurchaseOrdersOnFile"
            stLinkCriteria = "[ID]=" & nOrderID
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        Else
            If nOrderID = 0 Then
                 MsgBox " NO purchase order lines created."
            Else
                MsgBox "Error creating purchase order"
            End If
            DoCmd.Close acForm, Me.Name
        End If

    Case 2
        If DLookup("Value", "Settings", "KeyName = 'CreatePOBasedOnHistory_NewMethod'") Like "*True*" Then
            'CMB20250407 NEW METHOD
            'MsgBox "Sales History 1"
            If gcfHandleErrors Then On Error Resume Next
            'Create the table PurchaseOrderItemDaysSupply
            'because it is too 'complex' for Access
            DoCmd.DeleteObject acTable, "PurchaseOrderItemDaysSupply"
            If gcfHandleErrors Then On Error GoTo ErrLog
            datBegin = CDate(Forms![PurchaseOrderCreate]![txtBeginDate])
            datEnd = CDate(Forms![PurchaseOrderCreate]![txtEndDate])
            ndays = DateDiff("d", datBegin, datEnd) + 1

            With comQDF
                'qryPurchaseOrderItemReplenishMake
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryCreatePurchaseOrderBasedOnSalesHistory"
                .CommandType = adCmdStoredProc
                'ADO Numeric Datatypes are very particular
                .Parameters.Append .CreateParameter("parBeginDate", adDate, adParamInput)
                .Parameters.Append .CreateParameter("parEndDate", adDate, adParamInput)
                '.Parameters.Append .CreateParameter("parDaysSupply", adInteger, adParamInput)
                .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parBeginSalesDepartmentID", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parEndSalesDepartmentID", adBigInt, adParamInput)

                .Parameters("parBeginDate") = datBegin
                .Parameters("parEndDate") = datEnd
                '.Parameters("parDaysSupply") = nDays
                .Parameters("parVendorID") = CLng(Forms![PurchaseOrderCreate]![lstVendors])
                .Parameters("parBeginSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDBegin])
                .Parameters("parEndSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDEnd])
                .Execute
            End With

            Set comQDF = Nothing

            nOrderID = CreatePurchaseOrderBasedOnHistoryNewMethod(lstVendors, False, False)
            If nOrderID > 0 Then
                stDocName = "PurchaseOrdersOnFile"
                stLinkCriteria = "[ID]=" & nOrderID
                DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
                DoCmd.Close acForm, Me.Name
                Exit Sub
            Else
                If nOrderID = 0 Then
                     MsgBox " NO purchase order lines created."
                Else
                    MsgBox "Error creating purchase order"
                End If
                DoCmd.Close acForm, Me.Name
            End If
        Else
            'OLD METHOD

            'MsgBox "Sales History 1"
            If gcfHandleErrors Then On Error Resume Next
            'Create the table PurchaseOrderItemDaysSupply
            'because it is too 'complex' for Access
            DoCmd.DeleteObject acTable, "PurchaseOrderItemDaysSupply"
            If gcfHandleErrors Then On Error GoTo ErrLog
            datBegin = CDate(Forms![PurchaseOrderCreate]![txtBeginDate])
            datEnd = CDate(Forms![PurchaseOrderCreate]![txtEndDate])
            ndays = DateDiff("d", datBegin, datEnd) + 1

    '        Set qdfTemp = CurrentDb().QueryDefs("qryPurchaseOrderItemReplenishMake")
    '        qdfTemp.Parameters("parBeginDate") = datBegin
    '        qdfTemp.Parameters("parEndDate") = datEnd
    '        qdfTemp.Parameters("parVendorID") = CLng(Forms![PurchaseOrderCreate]![lstVendors])
    '        qdfTemp.Parameters("parBeginSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDBegin])
    '        qdfTemp.Parameters("parEndSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDEnd])
    '        qdfTemp.Execute
    '

            'MsgBox "Sales History 2"
            With comQDF
                'qryPurchaseOrderItemReplenishMake
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryPurchaseOrderItemReplenishMake"
                .CommandType = adCmdStoredProc
                'ADO Numeric Datatypes are very particular
                .Parameters.Append .CreateParameter("parBeginDate", adDate, adParamInput)
                .Parameters.Append .CreateParameter("parEndDate", adDate, adParamInput)
                '.Parameters.Append .CreateParameter("parDaysSupply", adInteger, adParamInput)
                .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parBeginSalesDepartmentID", adBigInt, adParamInput)
                .Parameters.Append .CreateParameter("parEndSalesDepartmentID", adBigInt, adParamInput)

                .Parameters("parBeginDate") = datBegin
                .Parameters("parEndDate") = datEnd
                '.Parameters("parDaysSupply") = nDays
                .Parameters("parVendorID") = CLng(Forms![PurchaseOrderCreate]![lstVendors])
                .Parameters("parBeginSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDBegin])
                .Parameters("parEndSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDEnd])
                .Execute
            End With

            Set comQDF = Nothing
            Set comQDF = CreateObject("ADODB.COMMAND")


            'added 10/1/2005
            'now add the 0 onhands with an order point set!

            'MsgBox "Sales History 3"
            With comQDF
                'qryPurchaseOrderItemReplenishMake
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryPurchaseOrderItemReplenishNoOnhandsAppend"
                .CommandType = adCmdStoredProc
                'ADO Numeric Datatypes are very particular
                .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
                .Parameters("parVendorID") = CLng(Forms![PurchaseOrderCreate]![lstVendors])
                .Execute
            End With

            Set comQDF = Nothing
            Set comQDF = CreateObject("ADODB.COMMAND")

            'MsgBox "Sales History 4"
            stDocName = "PurchaseOrderItemDaysSupplyReplenish"
            'DoCmd.OpenReport stDocName, acViewPreview

            ArchiveReport stDocName

            nOrderID = CreatePurchaseOrderDaysSupply(lstVendors, chkSaleItems.Value, chkTPR.Value)
            If nOrderID > 0 Then
                stDocName = "PurchaseOrdersOnFile"
                stLinkCriteria = "[ID]=" & nOrderID
                DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
                DoCmd.Close acForm, Me.Name
                Exit Sub
            Else
                If nOrderID = 0 Then
                     MsgBox " NO purchase order lines created."
                Else
                    MsgBox "Error creating purchase order"
                End If
                DoCmd.Close acForm, Me.Name
            End If
        End If

    Case 3 'all items supplied by this vendor
        stDocName = "PurchaseOrderCreate"
        ArchiveReport stDocName
        nOrderID = CreatePurchaseOrderItemsSupplied(lstVendors)
        If nOrderID > 0 Then
            stDocName = "PurchaseOrdersOnFile"
            stLinkCriteria = "[ID]=" & nOrderID
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        Else
            If nOrderID = 0 Then
                MsgBox " NO purchase order lines created."
            Else
                MsgBox "Error creating purchase order"
            End If
            DoCmd.Close acForm, Me.Name
        End If
    Case 4

        nOrderID = CreatePurchaseOrderBlankOrder(lstVendors)
        If nOrderID > 0 Then
            stDocName = "PurchaseOrdersOnFile"
            stLinkCriteria = "[ID]=" & nOrderID
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        Else
            If nOrderID = 0 Then
                MsgBox " NO purchase order lines created."
            Else
                MsgBox "Error creating purchase order"
            End If
            DoCmd.Close acForm, Me.Name
        End If
    Case 5
        If gcfHandleErrors Then On Error Resume Next
        'Create the table PurchaseOrderItemDaysSupply
        'because it is too 'complex' for Access
        DoCmd.DeleteObject acTable, "PurchaseOrderItemDaysSupply"
        If gcfHandleErrors Then On Error GoTo ErrLog
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryPurchaseOrderItemDaysSupplyMake"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parDaysSupply", adInteger, adParamInput)
            .Parameters.Append .CreateParameter("parBeginDate", adDate, adParamInput)
            .Parameters.Append .CreateParameter("parEndDate", adDate, adParamInput)
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            .Parameters.Append .CreateParameter("parBeginSalesDepartmentID", adBigInt, adParamInput)
            .Parameters.Append .CreateParameter("parEndSalesDepartmentID", adBigInt, adParamInput)

            .Parameters("parDaysSupply") = CInt(Forms![PurchaseOrderCreate]![txtDaysSupply])
            .Parameters("parBeginDate") = CDate(Forms![PurchaseOrderCreate]![txtBeginDate])
            .Parameters("parEndDate") = CDate(Forms![PurchaseOrderCreate]![txtEndDate])
            .Parameters("parVendorID") = CLng(Forms![PurchaseOrderCreate]![lstVendors])
            .Parameters("parBeginSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDBegin])
            .Parameters("parEndSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDEnd])
            .Execute
        End With

        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        'this will set all null UnitsPerOrder to 1
        stDocName = "PurchaseOrderItemDaysSupply"
        'DoCmd.OpenReport stDocName, acViewPreview

        ArchiveReport stDocName

        nOrderID = CreatePurchaseOrderDaysSupply(lstVendors, chkSaleItems.Value, chkTPR.Value)
        If nOrderID > 0 Then
            stDocName = "PurchaseOrdersOnFile"
            stLinkCriteria = "[ID]=" & nOrderID
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        Else
            If nOrderID = 0 Then
                 MsgBox " NO purchase order lines created."
            Else
                MsgBox "Error creating purchase order"
            End If
            DoCmd.Close acForm, Me.Name
        End If

    Case 6
        If gcfHandleErrors Then On Error Resume Next
        'Create the table PurchaseOrderItemWADS
        'because it is too 'complex' for Access
        DoCmd.DeleteObject acTable, "PurchaseOrderItemWADS"
        If gcfHandleErrors Then On Error GoTo ErrLog

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qyPurchaseOrderItemInfoByWADS"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
            '.Parameters.Append .CreateParameter("parBeginSalesDepartmentID", adBigInt, adParamInput)
            '.Parameters.Append .CreateParameter("parEndSalesDepartmentID", adBigInt, adParamInput)

            .Parameters("parVendorID") = CLng(Forms![PurchaseOrderCreate]![lstVendors])
            '.Parameters("parBeginSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDBegin])
            '.Parameters("parEndSalesDepartmentID") = CLng(Forms![PurchaseOrderCreate]![cboSalesDepartmentIDEnd])
            .Execute
        End With

        'round up the SOQ
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qyPurchaseOrderItemInfoByWADSAdjustSOQ"
            .CommandType = adCmdStoredProc
            .Execute
        End With

        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        stDocName = "PurchaseOrderItemsWADS"
        'DoCmd.OpenReport stDocName, acViewPreview

        ArchiveReport stDocName

        nOrderID = CreatePurchaseOrderWADS(lstVendors, chkSaleItems.Value, chkTPR.Value)
        If nOrderID > 0 Then
            stDocName = "PurchaseOrdersOnFile"
            stLinkCriteria = "[ID]=" & nOrderID
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        Else
            If nOrderID = 0 Then
                MsgBox " NO purchase order lines created."
            Else
                MsgBox "Error creating purchase order"
            End If
            DoCmd.Close acForm, Me.Name
        End If

    Case 7
        stDocName = "PurchaseOrderCreate"
        'DoCmd.OpenReport stDocName, acViewPreview
        ArchiveReport stDocName
        nOrderID = CreatePurchaseOrder(lstVendors, chkSaleItems.Value, chkTPR.Value)
        If nOrderID > 0 Then
            stDocName = "PurchaseOrdersOnFile"
            stLinkCriteria = "[ID]=" & nOrderID
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PurchaseOrderCreate"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        Else
            If nOrderID = 0 Then
                 MsgBox " NO purchase order lines created."
            Else
                MsgBox "Error creating purchase order"
            End If
            DoCmd.Close acForm, Me.Name
        End If
    Case Else
        MsgBox optBasePO & " case not found!"
        Exit Sub


    End Select
'    nOrderID = CreateClubOrder(cmbClubID, OG)
'    stDocName = "ClubOrder"
'    stLinkCriteria = "[ID]=" & nOrderID
'    DoCmd.Close
'    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Command4_Click"
    Resume ExitSub
End Sub



Private Sub cmdEditThisOrder_Click()
    Dim nOrderID As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
If gcfHandleErrors Then On Error GoTo ErrLog
    nOrderID = lstInvoices
    lbRetr.Caption = "Retrieving Order #" & nOrderID
    lbRetr.Visible = True
    Me.Repaint

    stDocName = "PurchaseOrdersOnFile"
    stLinkCriteria = "[ID]=" & nOrderID

    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
    DoCmd.Close acForm, Me.Name
ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 94
        MsgBox "Please select a Purchase order from the Green List before choosing Edit!"
    Case Else
        MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    End Select
    Resume ExitSub
End Sub

Private Sub cmdReport_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdReport_Click

    Dim stDocName As String

    Select Case optBasePO
    Case 1
        stDocName = "PurchaseOrderCreate"
    Case 2
    Case 3
    Case Else
        MsgBox optBasePO & " case not found!"
        Exit Sub
    End Select

        DoCmd.OpenReport stDocName, acPreview
        ArchiveReport stDocName


Exit_cmdReport_Click:
    Exit Sub

Err_cmdReport_Click:
    MsgBox Err.Description
    Resume Exit_cmdReport_Click

End Sub




Private Sub cmdOLCCEPSFTP_Click()
    'build the ftpeps.scr
    Dim strFZDir As String
    Dim strFZExe As String
    Dim stAppName As String

    MakeFTPScriptFiles
    'download files from olas
    If chkFTPFZActive Then
        strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files\FileZilla FTP Client\") 'false
        If strFZDir = "C:\Program Files\FileZilla FTP Client\" Then
            SetKMDSSettings "FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\"
        End If
        strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\") 'false
        strFZExe = GetKMDSSettings("FTPFZExe", "FileZilla.exe") 'false
        strFZDir = Trim(strFZDir)
        strFZExe = strFZExe
        'stAppName = strFZDir & strFZExe & " -c " & Chr(34) & "0/OLCC FTP" & Chr(34)
        stAppName = strFZDir & strFZExe & " --site=" & Chr(34) & "0/OLCC FTP"
        stAppName = stAppName & "/Electronic Packing List" & Chr(34)
        'MsgBox "Using Filezilla to receive Electronic Packing List" & vbCrLf & vbCrLf & "Choose 'Bookmarks -> Electronic Packing List" & vbCrLf & "Then double-click on most recent PackingList on the right"
        MsgBox "Using Filezilla to receive Electronic Packing List" & vbCrLf & vbCrLf & "Double-click on most recent PackingList on the right"
        'Call Shell(stAppName, 1)
        WaitForExecCmd (stAppName)
    Else
        If GetKMDSSettings("UseWinSCPForOLASTransmits", "False") = "False" Then
            WaitForExecCmd ("c:\npos\ftpeps.bat")
        Else ' use WINSCP
            '1. build .txt AND .bat files
            CreateWinSCPtoOLASFTPScriptFile ("EPS")
            '.2  execute .bat file.
            WaitForExecCmd ("c:\npos\winSCPEPS.bat")
        End If
    End If
    'import downloaded files
    OLCCElectronicPackingSlipImport
End Sub

Private Sub cmdOLCCEPSImport_Click()
Dim stDocName As String
stDocName = "OLCCPackingList"
    If gcfHandleErrors Then On Error Resume Next
    DoCmd.OpenForm stDocName, , , , , acDialog

    lstInvoices.Requery


    Me.Refresh

End Sub

Private Sub cmIDFOS_Click()
DoCmd.OpenForm "ItemDetailForOrderingFromOLAS"

End Sub

Private Sub dtpDateRangeFrom_Updated(Code As Integer)
FormatDateRanges
End Sub

Private Sub dtpDateRangeTo_Updated(Code As Integer)
FormatDateRanges
End Sub

Private Sub Form_Load()


'Dim RS As New ADODB.Recordset
Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")



Dim lngRecordsAffected As Long
Dim strSQL As String
Dim bFix As Boolean
Dim varReturn As Variant
    Dim strYTLink As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False


    'This will 'reprice' all orders with vendor #1
    'to fix the price for orders input in the previous month

    varReturn = SysCmd(acSysCmdSetStatus, "Delete PurchaseOrderLineUpdatePrice) ")
    On Error Resume Next

    DoCmd.DeleteObject acTable, "PurchaseOrderLineUpdatePrice"

    If gcfHandleErrors Then On Error GoTo ErrLog
    varReturn = SysCmd(acSysCmdSetStatus, "Create PurchaseOrderLineUpdatePrice ")
    strSQL = "qryPurchaseOrderLineUpdatePriceMakeTable"
    DoCmd.OpenQuery strSQL, acNormal, acEdit



    If gcfHandleErrors Then On Error GoTo ErrLog
    varReturn = SysCmd(acSysCmdSetStatus, "Reprice Orders with current Pricing (qryPurchaseOrderLineUpdatePrice) ")


    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderLineUpdatePrice"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    'to fix liquor PO's cost created in prior month
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPurchaseOrderLineUpdateCost"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")



    'to fix the 0 or NULL UnitsPerCase
    'are there any to fix? to prevent the 'catastrophic failure error
    varReturn = SysCmd(acSysCmdSetStatus, "Checking Units per Case")

    bFix = False
    strSQL = "SELECT Item.UnitsPerCase "
    strSQL = strSQL & "FROM Item "
    strSQL = strSQL & "WHERE (((Item.UnitsPerCase) Is Null)) OR "
    strSQL = strSQL & "(((Item.UnitsPerCase)<1))"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
        Set rs = comQDF.Execute
        If rs.BOF And rs.EOF Then
            bFix = False
        Else
            bFix = True
        End If
        rs.Close
        Set rs = Nothing
        Set rs = CreateObject("ADODB.Recordset")

    End With
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")


    If bFix Then
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryItemsSetUnitsPerCase"
            .CommandType = adCmdStoredProc
            comQDF.Execute 'lngRecordsAffected, , adCmdStoredProc
        End With
        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")
    End If

    'to fix the 0 or NULL UnitsPerOrder
    'are there any to fix? to prevent the 'catastrophic failure error
    varReturn = SysCmd(acSysCmdSetStatus, "Checking Units per Order")
    bFix = False
    strSQL = "SELECT Item.UnitsPerOrder "
    strSQL = strSQL & "FROM Item "
    strSQL = strSQL & "WHERE (((Item.UnitsPerOrder) Is Null)) OR "
    strSQL = strSQL & "(((Item.UnitsPerOrder)<1))"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
        Set rs = comQDF.Execute
        If rs.BOF And rs.EOF Then
            bFix = False
        Else
            bFix = True
        End If
        rs.Close
        Set rs = Nothing
        Set rs = CreateObject("ADODB.Recordset")
    End With
    Set comQDF = Nothing

    If bFix Then
        Set comQDF = CreateObject("ADODB.COMMAND")
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryItemsSetUnitsPerOrder"
            .CommandType = adCmdStoredProc
            .Execute
        End With
        Set comQDF = Nothing
    End If

    'are there any TPR's if so make the button visible
    varReturn = SysCmd(acSysCmdSetStatus, "Checking for TPRs")
    mbTPR = False
    strSQL = "SELECT Item.TPR "
    strSQL = strSQL & "FROM Item "
    strSQL = strSQL & "WHERE (((Item.TPR)<>0))"
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
        Set rs = comQDF.Execute
        If rs.BOF And rs.EOF Then
            mbTPR = False
        Else
            mbTPR = True
        End If
        rs.Close
        Set rs = Nothing
    End With
    Set comQDF = Nothing

    varReturn = SysCmd(acSysCmdSetStatus, "FINISHED")
    dtpDateRangeTo.Visible = False
    dtpDateRangeFrom.Visible = False
    lblDateRangeTo.Visible = False
    lblDateRangeFrom.Visible = False
    lblDateRange.Visible = False
    boxDateRange.Visible = False
    chkSaleItems.Visible = False
    chkTPR.Visible = False
    txtDaysSupply.Visible = False
    optBasePO.Visible = False
    cmdCreatePO.Visible = False
    cmdTPRClear.Visible = mbTPR
    lblVendorADS.Visible = False

    lstInvoices.Requery
    varReturn = SysCmd(acSysCmdSetStatus, "FormatDateRanges")
    FormatDateRanges
    varReturn = SysCmd(acSysCmdSetStatus, "FINISHED FormatDateRanges")

    'for future reference
    'this stored procedure set Items.AvgTurns to the day's supply
    'SQLPassThru ("sp_ItemDaysSupply")

'    If CurrentProject.AllForms("ScanBatch").IsLoaded Then
'        DoCmd.Close acForm, "ScanBatch", acSavePrompt
'    End If
    chkFTPFZActive = GetKMDSSettings("FTPFZActive", "0") 'false
'    added 31415
    chkUseWinSCPForOLASTransmits = GetKMDSSettings("UseWinSCPForOLASTransmits", "False") 'false

'    strYTLink = GetKMDSSettings("YouTubeRecElPackStep1", "https://youtu.be/9jQebGHGvlk")
'    If Left(strYTLink, 12) = "https://yout" Then
'        cmdYTRecElPackStep1.HyperlinkAddress = strYTLink
'        cmdYTRecElPackStep1.Visible = True
'    Else
'        cmdYTRecElPackStep1.Visible = False
'    End If
'    strYTLink = GetKMDSSettings("YouTubeRecElPackStep2", "https://youtu.be/gyHvHCfaLo4")
'    If Left(strYTLink, 12) = "https://yout" Then
'        cmdYTRecElPackStep2.HyperlinkAddress = strYTLink
'        cmdYTRecElPackStep2.Visible = True
'    Else
'        cmdYTRecElPackStep2.Visible = False
'    End If

    ' Item Detail For Ordering Subscription"
    strSQL = GetKMDSSettings("IDFOSubscription", "False")
    If UCase(strSQL) = "TRUE" Then
        cmIDFOS.Visible = True
    Else
        cmIDFOS.Visible = False
    End If



ExitSub:
    SetOptions True
    varReturn = SysCmd(acSysCmdClearStatus)
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub


Private Sub Frame74_AfterUpdate()

End Sub

Private Sub framPOSortBy_AfterUpdate()
    Dim strSQL As String

    Select Case framPOSortBy.Value
    Case 1
        strSQL = "SELECT PurchaseOrder.ID, [Vendor].[Name] & ' ' & [EPS] AS Vname, Format([OrderDate],'mm/dd') AS [Date], PurchaseOrder.id AS OrderNumber, qryPurchaseOrderLineTotals.Quantity, qryPurchaseOrderLineTotals.Dollars, Vendor.Name, IIf([PurchaseOrder.SetupDate]=#1/1/2018#,'(IMPORTED)','') AS EPS FROM (PurchaseOrder INNER JOIN Vendor ON PurchaseOrder.VendorID = Vendor.ID) INNER JOIN qryPurchaseOrderLineTotals ON PurchaseOrder.ID = qryPurchaseOrderLineTotals.PurchaseOrderID ORDER BY qryPurchaseOrderLineTotals.PurchaseOrderID, Vendor.Name"
    Case 2
        strSQL = "SELECT PurchaseOrder.ID, [Vendor].[Name] & ' ' & [EPS] AS Vname, Format([OrderDate],'mm/dd') AS [Date], PurchaseOrder.id AS OrderNumber, qryPurchaseOrderLineTotals.Quantity, qryPurchaseOrderLineTotals.Dollars, Vendor.Name, IIf([PurchaseOrder.SetupDate]=#1/1/2018#,'(IMPORTED)','') AS EPS FROM (PurchaseOrder INNER JOIN Vendor ON PurchaseOrder.VendorID = Vendor.ID) INNER JOIN qryPurchaseOrderLineTotals ON PurchaseOrder.ID = qryPurchaseOrderLineTotals.PurchaseOrderID ORDER BY  Vendor.Name, qryPurchaseOrderLineTotals.PurchaseOrderID"
    Case Else
        MsgBox "Unhandled case..."
        Exit Sub
    End Select

    With lstInvoices
        .RowSource = strSQL
        .Requery
    End With
End Sub

'Private Sub Form_Timer()
''    With lbRetr
''        .BackColor = (IIf(.BackColor = 0, 255, 0))
''    End With
''    Me.Repaint
'
'End Sub

Private Sub lstInvoices_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedInvoice = lstInvoices.Column(0) & " " & lstInvoices.Column(1)
    'cmdUseOrderGuide.Visible = Not IsNull(lstInvoices)
    'cmdNoOrderGuide.Visible = Not IsNull(lstInvoices)
    txtSelectedInvoice.Visible = Not IsNull(lstInvoices)
    cmdEditThisOrder.Visible = Not IsNull(lstInvoices)
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub



Private Sub lstInvoices_DblClick(Cancel As Integer)
cmdEditThisOrder_Click
End Sub

Private Sub lstInvoices_GotFocus()
Label0.Caption = "Edit Purchase Order"
End Sub

Private Sub lstVendors_AfterUpdate()
Dim buildstr As String

If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedVendor = lstVendors.Column(0) & " " & lstVendors.Column(1)
    txtSelectedVendor.Visible = Not IsNull(lstVendors)
    cmdCreatePO.Visible = Not IsNull(lstVendors)
    optBasePO.Visible = Not IsNull(lstVendors)
    chkSaleItems.Visible = Not IsNull(lstVendors)
    chkTPR.Visible = Not IsNull(lstVendors)
    If chkTPR.Visible Then chkTPR.Visible = mbTPR
    If Not gcfHandleErrors Then Debug.Print lstVendors.Column(2)
    cmbWADSDesc = lstVendors.Column(2)
    If IsNull(cmbWADSDesc.Column(0)) Or lstVendors.Column(3) <> 6 Then
        cmbWADSDesc.Visible = False
        optDaysSupplyWeighted.Visible = False
        lblVendorADS.Visible = True
    Else
        cmbWADSDesc.Visible = True
        optDaysSupplyWeighted.Visible = True
        lblVendorADS.Visible = False
    End If
    optBasePO = lstVendors.Column(3)

    Label49.Caption = lstVendors.Column(4)
    If Not IsNull(Trim(lstVendors.Column(1))) Or Len(Trim(lstVendors.Column(4))) > 10 Then
        buildstr = "http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q="
        buildstr = buildstr & lstVendors.Column(4)
        buildstr = buildstr & "&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
        buildstr = Replace(buildstr, " ", "+")
        lblMapQuest.Caption = "Click here for &Google Map to " & lstVendors.Column(1)
    Else
        lblMapQuest.Caption = "Click here for &Google Map"
        buildstr = "http://maps.google.com/maps"
    End If
    lblMapQuest.HyperlinkAddress = buildstr
    lblMapQuest.ControlTipText = buildstr


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "lstVendors_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub lstVendors_DblClick(Cancel As Integer)
    cmdCreatePO_Click
End Sub

Private Sub lstVendors_GotFocus()
Label0.Caption = "Create Purchase Order"

End Sub

Private Sub optBasePO_AfterUpdate()
    If optBasePO.Value = 5 Then
        chkIncludeTransferOut.Visible = True
    Else
        chkIncludeTransferOut.Visible = False
    End If


End Sub

Private Sub optBasePO_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    dtpDateRangeTo.Visible = False
    dtpDateRangeFrom.Visible = False
    lblDateRangeTo.Visible = False
    lblDateRangeFrom.Visible = False
    lblDateRange.Visible = False
    boxDateRange.Visible = False
    chkSaleItems.Visible = False
    txtDaysSupply.Visible = False
    chkTPR.Visible = False
    cboSalesDepartmentIDBegin.Visible = False
    cboSalesDepartmentIDEnd.Visible = False

    Select Case optBasePO
    Case 1 'order points
        chkSaleItems.Visible = True
        chkTPR.Visible = mbTPR

    Case 2  'sales history
        dtpDateRangeTo.Visible = True
        dtpDateRangeFrom.Visible = True
        lblDateRangeTo.Visible = True
        lblDateRangeFrom.Visible = True
        lblDateRange.Visible = True
        boxDateRange.Visible = True
        dtpDateRangeFrom = DateAdd("d", -7, Now)
'        dtpDateRangeFrom = CDate(DateAdd("d", -7, Now))
        dtpDateRangeTo = Now
        cboSalesDepartmentIDBegin.Visible = True
        cboSalesDepartmentIDEnd.Visible = True

    Case 5  'days supply
        dtpDateRangeTo.Visible = True
        dtpDateRangeFrom.Visible = True
        lblDateRangeTo.Visible = True
        lblDateRangeFrom.Visible = True
        lblDateRange.Visible = True
        boxDateRange.Visible = True
        dtpDateRangeFrom = CDate(DateAdd("m", -1, Now))
        dtpDateRangeTo = Now - 1
        txtDaysSupply.Visible = True
        chkSaleItems.Visible = True
        chkTPR.Visible = mbTPR
        cboSalesDepartmentIDBegin.Visible = True
        cboSalesDepartmentIDEnd.Visible = True

    Case 6  'WADS days supply
        chkSaleItems.Visible = True
        chkTPR.Visible = mbTPR
        cboSalesDepartmentIDBegin.Visible = True
        cboSalesDepartmentIDEnd.Visible = True
    Case Else
    End Select
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "optBasePO_Click"
    Resume ExitSub
End Sub
Private Sub Command22_Click()
If gcfHandleErrors Then On Error GoTo Err_Command22_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "PurchaseOrdersOnFile"

    stLinkCriteria = "[ID]=" & Me![lstVendors]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command22_Click:
    Exit Sub

Err_Command22_Click:
    MsgBox Err.Description
    Resume Exit_Command22_Click

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

Private Sub FormatDateRanges()
    txtBeginDate = Format(dtpDateRangeFrom, "mm/dd/yyyy 00:00")
    txtEndDate = Format(dtpDateRangeTo, "mm/dd/yyyy 23:59")

End Sub



Private Sub txtDaysSupply_Enter()
FormatDateRanges
End Sub
Private Sub cmdTPRClear_Click()

'Dim RS As New ADODB.Recordset
Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim lngRecordsAffected As Long
Dim strSQL As String
Dim bFix As Boolean

'    If gcfHandleErrors Then On Error GoTo ErrLog

    'to reset TPR to 0
    'are there any to fix? to prevent the 'catastrophic failure error
If gcfHandleErrors Then On Error GoTo Err_cmdTPRClear_Click
    If MsgBox("This will clear all of the TPR flags in the Item File" & vbCrLf & "Do you want to continue?", vbYesNo) = vbNo Then
        Exit Sub
    End If

    bFix = False
    strSQL = "SELECT Item.TPR "
    strSQL = strSQL & "FROM Item "
    strSQL = strSQL & "WHERE (((Item.TPR)<>0))"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = strSQL
        Set rs = comQDF.Execute
        If rs.BOF And rs.EOF Then
            bFix = False
        Else
            bFix = True
        End If
        rs.Close
        Set rs = Nothing
    End With
    Set comQDF = Nothing



    'This will 'reprice' all orders with vendor #1
    'to fix the price for orders input in the previous month
    If bFix = True Then
        CurrentDb().Execute "qryItemsSetTPRto0"
'        With comQDF
'            .ActiveConnection = CurrentProject.Connection
'            .CommandText = "qryItemsSetTPRto0"
'            .CommandType = adCmdStoredProc
'            .Execute
'        End With
'        Set comQDF = Nothing
    End If

Exit_cmdTPRClear_Click:
    mbTPR = False
    lstVendors.SetFocus
    cmdTPRClear.Visible = False
    Exit Sub

Err_cmdTPRClear_Click:
    MsgBox Err.Number & " " & Err.Description
    Resume Exit_cmdTPRClear_Click

End Sub
Private Sub cmdItemTPR_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdItemTPR_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ItemsTPR"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdItemTPR_Click:
    Exit Sub

Err_cmdItemTPR_Click:
    MsgBox Err.Description
    Resume Exit_cmdItemTPR_Click

End Sub
'------------------------------------------------------------
' cmdOLCCItemsNotOnFile_Click
'
'------------------------------------------------------------
Private Sub cmdOLCCItemsNotOnFile_Click()
On Error GoTo cmdOLCCItemsNotOnFile_Click_Err

    DoCmd.OpenForm "OLCCItemsNotOnFile", acNormal, "", "", , acDialog
    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="cmdOLCCItemsNotOnFile" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009
    ' _AXL:/11/forms"><Statements><Action Name="OpenForm"><Argument Name="FormName">OLCCItemsNotOnFile</Argument><Argument Name="WindowMode">Dialog</Argument></Action></Statements></UserInterfaceMacro>

    DoCmd.Close acForm, "PurchaseOrderCreate"
cmdOLCCItemsNotOnFile_Click_Exit:
    Exit Sub

cmdOLCCItemsNotOnFile_Click_Err:
    MsgBox Error$
    Resume cmdOLCCItemsNotOnFile_Click_Exit

End Sub


'------------------------------------------------------------
' cmdComparePurchaseOrders_Click
'
'------------------------------------------------------------
Private Sub cmdComparePurchaseOrders_Click()
On Error GoTo cmdComparePurchaseOrders_Click_Err

    DoCmd.OpenForm "PurchaseOrderCompare", acNormal, "", "", , acNormal


cmdComparePurchaseOrders_Click_Exit:
    Exit Sub

cmdComparePurchaseOrders_Click_Err:
    MsgBox Error$
    Resume cmdComparePurchaseOrders_Click_Exit

End Sub

Public Function CreatePurchaseOrderBasedOnHistoryNewMethod(VendorID As Long, IncludeSaleItems As Boolean, IncludeTPR As Boolean) As Integer
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

    CreatePurchaseOrderBasedOnHistoryNewMethod = 0

    'ArchiveReport "PurchaseOrderCreate"

    'strSQL = "SELECT * FROM PurchaseOrderItemDaysSupply "
    'strSQL = strSQL & "WHERE (((SOQ)>0))"

''    strSQL = "SELECT PurchaseOrderItemDaysSupply.* "
''    strSQL = strSQL & "FROM (PurchaseOrderItemDaysSupply "
''    strSQL = strSQL & "INNER JOIN Item ON PurchaseOrderItemDaysSupply.ItemId = Item.ID) "
''    strSQL = strSQL & "INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID "
''    strSQL = strSQL & "WHERE (((SubCategory.Description)<>'OBSOLETE') "
''    strSQL = strSQL & "AND ((PurchaseOrderItemDaysSupply.SOQ)>0))"

    strSQL = "SELECT * FROM PurchaseOrderItemDaysSupply"

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
            nLineCount = nLineCount + 1
            rsPOLine.AddNew
            rsPOLine!VendorID = VendorID
            rsPOLine!PurchaseOrderID = nPurchaseOrderID
            rsPOLine!LineID = nLineCount
            rsPOLine!ItemID = !ItemID
            rsPOLine!QuantityOrdered = !QuantityOrdered  '!SOQ
            rsPOLine!QuantityShipped = !QuantityShipped  '!SOQ
            rsPOLine!Price = !Price  '!Sprice
            rsPOLine!Cost = !ItemCost    '!Cost
            rsPOLine!Repack = !Repack

            rsPOLine!SetupDate = dtSetup
            rsPOLine!ChangeDate = dtSetup
            rsPOLine.Update

            .MoveNext
        Loop
        .Close
    End With
    Set rstemp = Nothing
''    Set rstemp = CreateObject("ADODB.Recordset")
''
''    If IncludeSaleItems = True Then
''        With comQDF
''            .ActiveConnection = CurrentProject.Connection
''            .CommandText = "qryPurchaseOrderCreateSaleItems"
''            .CommandType = adCmdStoredProc
''            'ADO Numeric Datatypes are very particular
''
'''            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
'''            .Parameters("parPurchaseOrderID") = nPurchaseOrderID
''            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
''            .Parameters("parVendorID") = VendorID
''        End With
''
''        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
''
''        With rstemp
''            If .RecordCount > 0 Then
''
''                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
''                'GoTo ErrLogShowRstemp
''                Do While Not .EOF
''                    nLineCount = nLineCount + 1
''                    rsPOLine.AddNew
''                    rsPOLine!VendorID = VendorID
''                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
''                    rsPOLine!LineID = nLineCount
''                    rsPOLine!ItemID = !ItemID
''                    rsPOLine!QuantityOrdered = 0
''                    rsPOLine!QuantityShipped = 0
''                    rsPOLine!Price = !SalePrice
''                    rsPOLine!Cost = !Cost
''
''                    rsPOLine!SetupDate = dtSetup
''                    rsPOLine!ChangeDate = dtSetup
''                    rsPOLine.Update
''                    If Not gcfHandleErrors Then Debug.Print !Code; " "; !SOQ
''                    .MoveNext
''                Loop
''            End If
''            .Close
''        End With
''        Set rstemp = Nothing
''        Set rstemp = CreateObject("ADODB.Recordset")
''        Set comQDF = Nothing
''        Set comQDF = CreateObject("ADODB.COMMAND")
''    End If
''
''    If IncludeTPR = True Then
''        With comQDF
''            .ActiveConnection = CurrentProject.Connection
''            .CommandText = "qryPurchaseOrderCreateTPRItems"
''            .CommandType = adCmdStoredProc
''            'ADO Numeric Datatypes are very particular
''
'''            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
'''            .Parameters("parPurchaseOrderID") = nPurchaseOrderID
''            .Parameters.Append .CreateParameter("parVendorID", adBigInt, adParamInput)
''            .Parameters("parVendorID") = VendorID
''        End With
''
''        rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
''
''        With rstemp
''            If .RecordCount > 0 Then
''
''                If gcfHandleErrors Then On Error Resume Next 'blast past the dupes
''                'GoTo ErrLogShowRstemp
''                Do While Not .EOF
''                    nLineCount = nLineCount + 1
''                    rsPOLine.AddNew
''                    rsPOLine!VendorID = VendorID
''                    rsPOLine!PurchaseOrderID = nPurchaseOrderID
''                    rsPOLine!LineID = nLineCount
''                    rsPOLine!ItemID = !ItemID
''                    rsPOLine!QuantityOrdered = 0
''                    rsPOLine!QuantityShipped = 0
''                    rsPOLine!Price = !Price
''                    rsPOLine!Cost = !Cost
''
''                    rsPOLine!SetupDate = dtSetup
''                    rsPOLine!ChangeDate = dtSetup
''                    rsPOLine.Update
''                    If Not gcfHandleErrors Then Debug.Print !ItemID; " "; !SOQ
''                    .MoveNext
''                Loop
''            End If
''            .Close
''        End With
''        Set rstemp = Nothing
''        Set rstemp = CreateObject("ADODB.Recordset")
''        Set comQDF = Nothing
''        Set comQDF = CreateObject("ADODB.COMMAND")
''
''    End If

ExitSub:
    CreatePurchaseOrderBasedOnHistoryNewMethod = nPurchaseOrderID
    Exit Function

ErrLogShowRstemp:
    strMsg = "Error when saving item code: " & rstemp!Code
ErrLog:
    'strMSG = "Error when saving item code: " & rsTemp!ItemID

    MsgBox strMsg & vbCrLf & Err.Number & " " & Err.Description
    nLineCount = -1
    Resume ExitSub
End Function
