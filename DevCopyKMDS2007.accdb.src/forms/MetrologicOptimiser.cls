Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'------------------------------------------------------------
' Command0_Click
'
'------------------------------------------------------------

Dim mstrFileName As String
Dim mstrBatchFileName As String


Private Sub cboCategoryIDBegin_AfterUpdate()
    UpdateSettings
End Sub


Private Sub cboCategoryIDEnd_AfterUpdate()
    UpdateSettings
End Sub

Private Sub cboNotScannedCategoryIDBegin_AfterUpdate()
    If cboNotScannedCategoryIDBegin > cboNotScannedCategoryIDEnd Then cboNotScannedCategoryIDEnd = cboNotScannedCategoryIDBegin
    MetrologicOptimiserSubItemsNotScanned.Requery

End Sub

Private Sub cboNotScannedCategoryIDEnd_AfterUpdate()
    If cboNotScannedCategoryIDEnd < cboNotScannedCategoryIDBegin Then cboNotScannedCategoryIDBegin = cboNotScannedCategoryIDEnd
    MetrologicOptimiserSubItemsNotScanned.Requery
End Sub

Private Sub cboSalesDepartmentIDBegin_AfterUpdate()
    UpdateSettings
End Sub

Private Sub cboSalesDepartmentIDBegin_Click()
    UpdateSettings
End Sub

Private Sub cboSalesDepartmentIDEnd_AfterUpdate()
    UpdateSettings
End Sub


Private Sub cmdContinue_Click()
    Dim nID As Long
    Dim strName As String
    Dim bTransferIn As Boolean
    Dim strInOrOut As String
    Dim lngTransferID As Long

    With lstVendorClubAgent
        If .ItemData(.ListIndex + 1) = "ID" Then
            MsgBox "First select a Name from the list and then click the 'Continue' button"
        Else
            nID = .ItemData(.ListIndex + 1)
            strName = .Column(0, .ListIndex + 1)
            Select Case .Tag
            Case "Vendor"
                If MsgBox("Create a PURCHASE ORDER for Vendor #" & nID & " " & strName & "?", vbYesNo) = vbYes Then
                    If CreatePurchaseOrderFromMetroLogic(nID) <> 0 Then
                        DoCmd.Close
                        DoCmd.OpenForm "PurchaseOrderCreate"
                    End If
                End If
            Case "Club"
                If MsgBox("Create a LICENSEE ORDER for Club #" & nID & " " & strName & "?", vbYesNo) = vbYes Then
                    If CreateClubOrderFromMetroLogic(nID) > 0 Then
                        DoCmd.Close
                        DoCmd.OpenForm "ClubOrderCreate"
                    End If
                End If
            Case "Agent"
                If MsgBox("Create a TRANSFER for Agent #" & nID & " " & strName & "?", vbYesNo) = vbYes Then
                    If MsgBox("Is this a Transfer Out? (Leaving your store?)", vbYesNo) = vbYes Then
                        bTransferIn = False
                        strInOrOut = "O"
                    Else
                        bTransferIn = True
                        strInOrOut = "I"
                    End If
                    lngTransferID = CreateTransfer(nID, strInOrOut)
                    If CreateStoreTransferFromMetrologic(nID, lngTransferID, strInOrOut) > 0 Then
                        DoCmd.Close
                        If bTransferIn Then
                            DoCmd.OpenForm "AgentTransferIn", acNormal, "", "", , acNormal
                        Else
                            DoCmd.OpenForm "AgentTransferOut", acNormal, "", "", , acNormal
                        End If

                    End If
                End If
            Case "OLCCClaim"

                If MsgBox("Create a Claim for #" & nID & " " & strName & "?", vbYesNo) = vbYes Then
                    strInOrOut = "C" 'claim
                    lngTransferID = CreateTransfer(nID, strInOrOut)
                    If CreateStoreTransferFromMetrologic(nID, lngTransferID, strInOrOut) > 0 Then
                        DoCmd.Close
                        DoCmd.OpenForm "AgentTransferOutClaims", acNormal, "", "", , acNormal

                    End If
                End If


            Case Else
            End Select
        End If
    End With
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



Private Sub cmdPrintAuditLog_Click()
On Error GoTo cmdPrintAuditLog_Click_Err

    DoCmd.OpenReport "MetrologicItemsScanBatchAudit", acViewPreview, "", "", acNormal
    'DoCmd.RunCommand acCmdPrint


cmdPrintAuditLog_Click_Exit:
    Exit Sub

cmdPrintAuditLog_Click_Err:
    MsgBox Error$
    Resume cmdPrintAuditLog_Click_Exit
End Sub

Private Sub cmdPrintVariance_Click()
On Error GoTo cmdPrintVariance_Click_Err

    DoCmd.OpenReport "MetrologicItemsScanBatchVariance", acViewPreview, "", "", acNormal
    'DoCmd.RunCommand acCmdPrint


cmdPrintVariance_Click_Exit:
    Exit Sub

cmdPrintVariance_Click_Err:
    MsgBox Error$
    Resume cmdPrintVariance_Click_Exit

End Sub

Private Sub cmdReceive_Click()

Dim strSQL As String
Dim comQDF As Object
Dim comQDF2 As Object

'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fso As New FileSystemObject

Dim strSourceFname
Dim strDestFName As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    WaitForExecCmd "C:\kmds\metrologic\Program\Data_Read.exe"

    'archive the file before you do anything else..
    strSourceFname = "C:\KMDS\Metrologic\In\MetrologicScanBatch.txt"
    strDestFName = "C:\KMDS\Metrologic\In\Archive\" & Format(Now, "yyyy-mm-dd hh-nn-ss") & "MetrologicScanBatch.txt"
    DoCmd.TransferText acImportDelim, "MetrologicScanbatchImportSpec", "MetrologicScanBatch", "C:\KMDS\Metrologic\In\MetrologicScanBatch.txt", False

    fso.MoveFile strSourceFname, strDestFName

    'if user input of qty: is not numeric then make it "-1"
    '06/29/17 if user input of qty: is not numeric then make it "0"
    Set comQDF = CreateObject("ADODB.COMMAND")
    strSQL = "qryMetrologicScanBatch_Validate_Qty"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
    Set comQDF = Nothing

    'this is implementing the kmds scan card
    Set comQDF = CreateObject("ADODB.COMMAND")
    strSQL = "qryMetrologicScanBatch_Upd_Qty"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
    Set comQDF = Nothing

    Set comQDF2 = CreateObject("ADODB.COMMAND")
    strSQL = "qryMetrologicScanBatch_Upd_KMDSQty"
    With comQDF2
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
    Set comQDF2 = Nothing

    Set comQDF2 = CreateObject("ADODB.COMMAND")
    strSQL = "qryMetrologicScanBatch_Upd_CashierID"
    With comQDF2
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("parCashierID", adBigInt, adParamInput)
        .Parameters("parCashierID") = glnCashierID
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
    Set comQDF2 = Nothing

    Set comQDF2 = CreateObject("ADODB.COMMAND")
    strSQL = "qryMetrologicScanBatch_DeleteHeaderScans"
    With comQDF2
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
    Set comQDF2 = Nothing

    MetrologicOptimiserSubAuditLog.Requery
    MetrologicScanBatchSubVariance.Requery
    MetrologicOptimiserSubItemsNotScanned.Requery
    ShowFinishButtons
    MetrologicOptimiserSubAuditLog.SetFocus

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub

Private Sub cmdSend_Click()
Dim strFileName As String
Dim strSQL As String
Dim nRecordsToSend As Long

    If gcfHandleErrors Then On Error GoTo ErrLog

    AppendOLCCShelfTags
    nRecordsToSend = BuildLookupFile
    If nRecordsToSend = 0 Then
        MsgBox "Please choose a valid range of Sales Departments AND Categories"
    Else
        DownloadLookupFile
    End If

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub


End Sub


'SUBROUTINES

Sub AppendOLCCShelfTags()
'Dim comQDF As New Command
Dim comQDF As Object
Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    strSQL = "qryMetrologicForOLCCShelfTagsAppend"

    Set comQDF = CreateObject("ADODB.COMMAND")

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub

Public Function BuildLookupFile() As Long

Dim nOrderID As Long
Dim stDocName As String

Dim stLinkCriteria As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Long  'Integer
Dim mstrNow As String

Dim strFileName As String
Dim strSQL As String
Dim strCMD As String
Dim strHeader As String

Dim strDescWithNoVowels As String
Dim strShelf As String
Dim strPrice As String


    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")


    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic")
    End If
    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\In") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\In")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\In\Archive") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\In\Archive")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\Out") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\Out")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\Program") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\Program")
    End If

    strSQL = "qryMetrologicOptimiserLookup"
    mstrFileName = "C:\KMDS\Metrologic\Out\" & strSQL & ".txt"
    Set txtstrTemp = fsoTemp.OpenTextFile(mstrFileName, ForWriting, True, TristateFalse)


    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("prmSalesdepartmentIdBegin", adBigInt, adParamInput)
        .Parameters("prmSalesdepartmentIdBegin") = cboSalesDepartmentIDBegin
        .Parameters.Append .CreateParameter("prmSalesdepartmentIdEnd", adBigInt, adParamInput)
        .Parameters("prmSalesdepartmentIdEnd") = cboSalesDepartmentIDEnd

        .Parameters.Append .CreateParameter("prmCategoryDescBegin", adBSTR, adParamInput)
        .Parameters("prmCategoryDescBegin") = cboCategoryIDBegin
        .Parameters.Append .CreateParameter("prmCategoryDescEnd", adBSTR, adParamInput)
        .Parameters("prmCategoryDescEnd") = cboCategoryIDEnd
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If
        '"001292        LD GRND D 121292  22.95 "
        '"-1            Ver  12345  05/12 17:44 "
        'glnVersionNumber
        strHeader = "-1            Version" & Format(glnVersionNumber, "@@@@@") & Format(Now, "mm/dd hh:mm") & " "
        txtstrTemp.WriteLine strHeader

        Do While Not .EOF
'            If !SalesDepartmentID >= [Forms]![MetrologicOptimiser]![cboSalesDepartmentIDBegin] Then
'               If !SalesDepartmentID <= [Forms]![MetrologicOptimiser]![cboSalesDepartmentIDEnd] Then
'                    If !CatDesc >= [Forms]![MetrologicOptimiser]![cboCategoryIDBegin] Then
'                        If !CatDesc <= [Forms]![MetrologicOptimiser]![cboCategoryIDEnd] Then
                            nLineCount = nLineCount + 1
                            strDescWithNoVowels = UCase(!Description)

                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "A", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "E", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "I", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "O", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "U", "")
                            strDescWithNoVowels = Format(Left(strDescWithNoVowels, 9), "!@@@@@@@@@")
                            strShelf = Format(!shelf, "@@@")
                            If !shelf > 999 Then
                                strShelf = "999"
                            End If
                            If !shelf < -99 Then
                                strShelf = "-99"
                            End If
                            strPrice = !pri
                            If Len(strPrice) > 6 Then strPrice = Left(strPrice, 6)



                            txtstrTemp.WriteLine !UPC & strDescWithNoVowels & strShelf & !Item & strPrice  '& "$"
'                        End If
'                    End If
'                End If
'            End If
            .MoveNext
        Loop
    End With
    txtstrTemp.Close
    txtCount = nLineCount
    Me.Refresh
    MsgBox nLineCount & " records to send!"
    BuildLookupFile = nLineCount

'mstrBatchFileName = "C:\kmds\metrologic\Program\DlookupKMDS.bat"

'strCMD = "C:\kmds\metrologic\Program\Dlookup.exe "
'strCMD = strCMD & mstrFileName & " "
'strCMD = strCMD & GetKMDSSettings("MetrologicCradleCommPort", "6") & " "
'Set txtstrtemp = fsotemp.OpenTextFile(mstrBatchFileName, ForWriting, True, TristateFalse)
'txtstrtemp.WriteLine strCMD
'txtstrtemp.Close

ExitSub:
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Function

Public Function BuildLookupFileWithCaseQty() As Long

Dim nOrderID As Long
Dim stDocName As String

Dim stLinkCriteria As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Long
Dim mstrNow As String

Dim strFileName As String
Dim strSQL As String
Dim strCMD As String
Dim strHeader As String

Dim strDescWithNoVowels As String
Dim strShelf As String
Dim strPrice As String


    If gcfHandleErrors Then On Error GoTo ErrLog
'    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    SetOptions False

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic")
    End If
    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\In") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\In")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\In\Archive") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\In\Archive")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\Out") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\Out")
    End If

    If Not fsoTemp.FolderExists("C:\KMDS\Metrologic\Program") Then
        fsoTemp.CreateFolder ("C:\KMDS\Metrologic\Program")
    End If

    On Error Resume Next
    DoCmd.DeleteObject acTable, "MetrologicLookup"
    If gcfHandleErrors Then On Error GoTo ErrLog
    strSQL = "qryMetrologicLookup_01_MakeTable"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("prmSalesdepartmentIdBegin", adBigInt, adParamInput)
        .Parameters("prmSalesdepartmentIdBegin") = cboSalesDepartmentIDBegin
        .Parameters.Append .CreateParameter("prmSalesdepartmentIdEnd", adBigInt, adParamInput)
        .Parameters("prmSalesdepartmentIdEnd") = cboSalesDepartmentIDEnd

        .Parameters.Append .CreateParameter("prmCategoryDescBegin", adBSTR, adParamInput)
        .Parameters("prmCategoryDescBegin") = cboCategoryIDBegin
        .Parameters.Append .CreateParameter("prmCategoryDescEnd", adBSTR, adParamInput)
        .Parameters("prmCategoryDescEnd") = cboCategoryIDEnd
        'ADO Numeric Datatypes are very particular
        .Execute
    End With
    ' thankfully OLCC rethought this cluster

'    strsql = "qryMetroLogicLookup_02_Append_10_Digit_Shelf_Tags"
'    With comQDF
'        .ActiveConnection = CurrentProject.Connection
'        .CommandText = strsql
'        .CommandType = adCmdStoredProc
'        .Parameters.Append .CreateParameter("prmSalesdepartmentIdBegin", adBigInt, adParamInput)
'        .Parameters("prmSalesdepartmentIdBegin") = cboSalesDepartmentIDBegin
'        .Parameters.Append .CreateParameter("prmSalesdepartmentIdEnd", adBigInt, adParamInput)
'        .Parameters("prmSalesdepartmentIdEnd") = cboSalesDepartmentIDEnd
'
'        .Parameters.Append .CreateParameter("prmCategoryDescBegin", adBSTR, adParamInput)
'        .Parameters("prmCategoryDescBegin") = cboCategoryIDBegin
'        .Parameters.Append .CreateParameter("prmCategoryDescEnd", adBSTR, adParamInput)
'        .Parameters("prmCategoryDescEnd") = cboCategoryIDEnd
'        'ADO Numeric Datatypes are very particular
'        .Execute
'    End With

    strSQL = "qryMetrologicLookup_03"
    mstrFileName = "C:\KMDS\Metrologic\Out\qryMetrologicOptimiserLookup.txt"
    Set txtstrTemp = fsoTemp.OpenTextFile(mstrFileName, ForWriting, True, TristateFalse)


    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("prmSalesdepartmentIdBegin", adBigInt, adParamInput)
        .Parameters("prmSalesdepartmentIdBegin") = cboSalesDepartmentIDBegin
        .Parameters.Append .CreateParameter("prmSalesdepartmentIdEnd", adBigInt, adParamInput)
        .Parameters("prmSalesdepartmentIdEnd") = cboSalesDepartmentIDEnd

        .Parameters.Append .CreateParameter("prmCategoryDescBegin", adBSTR, adParamInput)
        .Parameters("prmCategoryDescBegin") = cboCategoryIDBegin
        .Parameters.Append .CreateParameter("prmCategoryDescEnd", adBSTR, adParamInput)
        .Parameters("prmCategoryDescEnd") = cboCategoryIDEnd
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        If rstemp.RecordCount = 0 Then
            MsgBox "No Records found for Sales department 1"
            GoTo ExitSub
        End If
        '"001292        LD GRND D 121292  22.95 "
        '"-1            Ver  12345  05/12 17:44 "
        'glnVersionNumber
        strHeader = "-1            Version" & Format(glnVersionNumber, "@@@@@") & Format(Now, "mm/dd hh:mm") & " 12345678"
        txtstrTemp.WriteLine strHeader

        Do While Not .EOF
'            If !SalesDepartmentID >= [Forms]![MetrologicOptimiser]![cboSalesDepartmentIDBegin] Then
'               If !SalesDepartmentID <= [Forms]![MetrologicOptimiser]![cboSalesDepartmentIDEnd] Then
'                    If !CatDesc >= [Forms]![MetrologicOptimiser]![cboCategoryIDBegin] Then
'                        If !CatDesc <= [Forms]![MetrologicOptimiser]![cboCategoryIDEnd] Then
                            nLineCount = nLineCount + 1
                            strDescWithNoVowels = UCase(!Description)

                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "A", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "E", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "I", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "O", "")
                            If Len(strDescWithNoVowels) > 9 Then strDescWithNoVowels = Replace(strDescWithNoVowels, "U", "")
                            strDescWithNoVowels = Format(Left(strDescWithNoVowels, 9), "!@@@@@@@@@")
                            strShelf = Format(!shelf, "@@@")
                            If !shelf > 999 Then
                                strShelf = "999"
                            End If
                            If !shelf < -99 Then
                                strShelf = "-99"
                            End If
                            strPrice = !pri
                            If Len(strPrice) > 6 Then strPrice = Left(strPrice, 6)



                            txtstrTemp.WriteLine !UPC & strDescWithNoVowels & strShelf & !Item & strPrice & !UCase '& "$"
'                        End If
'                    End If
'                End If
'            End If
            .MoveNext
        Loop
    End With
    txtstrTemp.Close
    txtCount = nLineCount
    Me.Refresh
    MsgBox nLineCount & " records to send!"
    BuildLookupFileWithCaseQty = nLineCount

'mstrBatchFileName = "C:\kmds\metrologic\Program\DlookupKMDS.bat"

'strCMD = "C:\kmds\metrologic\Program\Dlookup.exe "
'strCMD = strCMD & mstrFileName & " "
'strCMD = strCMD & GetKMDSSettings("MetrologicCradleCommPort", "6") & " "
'Set txtstrtemp = fsotemp.OpenTextFile(mstrBatchFileName, ForWriting, True, TristateFalse)
'txtstrtemp.WriteLine strCMD
'txtstrtemp.Close

ExitSub:
    SetOptions True
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Function


Public Function DownloadLookupFile()

' lookup.ini contains
'C:\KMDS\Metrologic\Out\qryMetrologicOptimiserLookup.txt
'6
'1
'2
'1

'help file for dlookup.exe
'Dlookup.exe supports the following command line arguments.
'
'Example:
'DLOOKUP LOOKUP.TXT,1,1 //COM1, 115200bps, RS-232
'DLOOKUP LOOKUP.TXT,2,1,2 //COM2, 115200bps, Cradle-IR
'If command line arguments are not given, user can select the appropriate parameters from the pop-up dialog box.
'
'DLookup [FileName], [COM port], [Baud rate], [Interface], [Show dialog]
'
'COM Port
'1 ~ 255
'Default: 1
'
'Baud Rate
'1 for 115200 bps
'2 for 57600 bps
'3 for 38400 bps
'4 for 19200 bps
'5 for 9600 bps
'Default: 1
'
'Interface
'1 for RS-232 or IrDA
'2 for Cradle-IR
'3 for MODEM
'Default: 1
'
'show dialog
'0 for Disabled (= hide dialog)
'1 for Enabled (= show dialog)
'Default: 0


'strCMD = strCMD & "1 " ' baud = 115200
'strCMD = strCMD & "2 " ' cradle
'strCMD = strCMD & "1 " ' show dialog

WaitForExecCmd "C:\kmds\metrologic\Program\Dlookup.exe"

ExitSub:
    Exit Function
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Function


Private Sub cmdSendCaseQty_Click()
Dim strFileName As String
Dim strSQL As String
Dim nRecordsToSend As Long

    If gcfHandleErrors Then On Error GoTo ErrLog

    ' thankfully OLCC rethought this cluster
    'AppendOLCCShelfTags
    nRecordsToSend = BuildLookupFileWithCaseQty
    If nRecordsToSend = 0 Then
        MsgBox "Please choose a valid range of Sales Departments AND Categories"
    Else
        DownloadLookupFile
    End If

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub

Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog
    cboSalesDepartmentIDBegin = GetKMDSSettings("MetrologicSalesDepartmentBegin", "1")
    cboSalesDepartmentIDEnd = GetKMDSSettings("MetrologicSalesDepartmentEnd", "1")
    cboCategoryIDBegin = GetKMDSSettings("MetrologicCategoryBegin", "ZZZ")

    If cboCategoryIDBegin = "Not Assigned" Then cboCategoryIDBegin = " Not Assigned"
    cboCategoryIDEnd = GetKMDSSettings("MetrologicCategoryEnd", "ZZZ")
    If cboCategoryIDEnd = "Not Assigned" Then cboCategoryIDEnd = " Not Assigned"



    txtSalesDepartmentBegin = cboSalesDepartmentIDBegin
    txtSalesDepartmentEnd = cboSalesDepartmentIDEnd
    lstVendorClubAgent.Visible = False
    cmdContinue.Visible = False

    cmdSendCaseQty.Visible = GetKMDSSettings("MetrologicSendCaseQty", "False")
    cmdSend.Visible = Not cmdSendCaseQty.Visible

    cboNotScannedCategoryIDBegin = cboCategoryIDBegin
    cboNotScannedCategoryIDEnd = cboCategoryIDEnd
    MetrologicOptimiserSubItemsNotScanned.Requery

    Me.InsideHeight = boxScreen.Height + 100
    Me.InsideWidth = boxScreen.Width + 500
    ShowFinishButtons

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    ShowColumns

End Sub

Private Sub framBatch_Click()
    Select Case framBatch.Value
    Case 1
        MovetoD7
    Case 2
        CreatePO
    Case 3
        ClubOrder
    Case 4
        StoreTransfer
    Case 5
        DeleteBatch
    Case 6
        OLCCClaims
    Case Else
        MsgBox "Select an option to continue..."
    End Select
End Sub

Private Sub lstVendorClubAgent_DblClick(Cancel As Integer)
    cmdContinue_Click
End Sub

Private Sub MetrologicOptimiserSubAuditLog_Exit(Cancel As Integer)
    UpdateSettings

End Sub

Private Sub MetrologicOptimiserSubItemsNotScanned_Enter()
    If gcfHandleErrors Then On Error GoTo ErrLog
    MetrologicOptimiserSubItemsNotScanned.Requery
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub


Private Sub MetrologicScanBatchSubVariance_Enter()
    If gcfHandleErrors Then On Error GoTo ErrLog
    MetrologicOptimiserSubAuditLog.Requery
    MetrologicScanBatchSubVariance.Requery
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub cmdPrintItemsNotScanned_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.OpenReport "MetrologicItemsNotScanned", acViewPreview, "", "", acNormal
    'DoCmd.RunCommand acCmdPrint


ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub

Private Sub pgItems_Click()
    If gcfHandleErrors Then On Error GoTo ErrLog
    MetrologicOptimiserSubAuditLog.Requery
    MetrologicScanBatchSubVariance.Requery
    MetrologicOptimiserSubItemsNotScanned.Requery
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub


Private Sub ClubOrder()
Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    lblVendorClubAgent.Caption = "Select a Club to create a licensee order"
    strSQL = "SELECT Name, ID FROM Club  ORDER BY Name"
    With lstVendorClubAgent
        .RowSource = strSQL
        .Requery
        .Tag = "Club"
        .BackColor = lblClubOrder.BackColor
        .Visible = True
    End With

    cmdContinue.Visible = True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub

Private Sub CreatePO()

Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    lblVendorClubAgent.Caption = "Select a Vendor to create a purchase order"
    strSQL = "SELECT Name, ID FROM Vendor WHERE (((Vendor.ID)>0)) ORDER BY Vendor.Name"
    With lstVendorClubAgent
        .RowSource = strSQL
        .Requery
        .Tag = "Vendor"
        .BackColor = lblCreatePO.BackColor
        .Visible = True
    End With

    cmdContinue.Visible = True

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub





Private Sub DeleteBatch()
Dim comQDF As Object
'Dim nAgentNumber As Long
Dim strSQL As String
    If gcfHandleErrors Then On Error GoTo ErrLog

    If MsgBox("This will delete all of the scanned records in the Audit Log, Are you sure?", vbYesNoCancel) = vbYes Then

        Set comQDF = CreateObject("ADODB.COMMAND")
        strSQL = "qryMetrologicScanBatchDelete"


        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = strSQL
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Execute
        End With

        MetrologicOptimiserSubAuditLog.Requery
        MetrologicScanBatchSubVariance.Requery
        MsgBox "Deleted!"
        If cmdSend.Visible Then
            cmdSend.SetFocus
        Else
            cmdSendCaseQty.SetFocus
        End If
        ShowFinishButtons
    Else
        MetrologicOptimiserSubAuditLog.SetFocus
    End If



ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub



Public Sub ExportMetrologicScanBatch()
Dim strOutFname As String
On Error Resume Next

MkDir ("C:\kmds\Metrologic\in\archive")
If gcfHandleErrors Then On Error GoTo ErrLog

DoCmd.TransferText acExportDelim, , "MetrologicScanBatch", "C:\kmds\Metrologic\in\archive\msb" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".txt", True
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description & " Export Scan Batch Line Counts", vbRed, True
    Resume ExitSub
End Sub




Private Sub MovetoD7()
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo ErrLog

    If MsgBox("Create a batch in the 'D-7 Audit Inventory' screen?", vbYesNo) = vbYes Then
        framBatch.Value = 0

        rstemp.Open "Audit", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        If rstemp.EOF And rstemp.BOF Then
        Else
            MsgBox "The D-7 Audit Inventory screen already has a batch pending; this batch must be posted or deleted before this is allowed!", vbCritical
            rstemp.Close
            framBatch.Value = 0
            Exit Sub
        End If

        SetOptions False
        ArchiveReport "MetrologicItemsScanBatchVariance"

        If rstemp.EOF And rstemp.BOF Then
            DoCmd.OpenQuery "qryMetrologicAppendToAudit", acViewNormal, acEdit
            DoCmd.OpenQuery "qryScanBatchAppendToAuditRecalcVariance", acViewNormal, acEdit

            '220113 DoCmd.OpenQuery "qryMetrologicAppendToInvoLine", acViewNormal, acEdit
            ExportMetrologicScanBatch
            DoCmd.RunSQL "Delete * from MetrologicScanBatch"
            DoCmd.Close acForm, "MetrologicOptimiser"
        End If
    End If
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub

Private Sub OLCCClaims()
On Error GoTo cmdStoreTransfer_Click_Err

    Dim strSQL As String
    lblVendorClubAgent.Caption = "Select where the Claim is going"


    strSQL = "SELECT  LastName as Name, AgentNumber as ID FROM Agents "
    strSQL = strSQL & "WHERE (((Agents.LastName) Like '*Claim*') And ((Agents.ActiveForTransfer) = True)) "
    strSQL = strSQL & "ORDER BY Agents.AgentNumber;"


    With lstVendorClubAgent
        .RowSource = strSQL
        .Requery
        .Tag = "OLCCClaim"
        .BackColor = lblOLCCClaim.BackColor
        .Visible = True
    End With
    cmdContinue.Visible = True

cmdStoreTransfer_Click_Exit:
    Exit Sub

cmdStoreTransfer_Click_Err:
    MsgBox Error$
    Resume cmdStoreTransfer_Click_Exit
End Sub




Private Sub ShowColumns()


    With [Forms]![MetrologicOptimiser]![MetrologicOptimiserSubAuditLog].[Form]

        !ScanTime.ColumnWidth = -2
        !ScanDate.ColumnWidth = -2
        !CurPrice.ColumnWidth = -2
        !Quantity.ColumnWidth = -2
        !ScanCode.ColumnWidth = -2
        !CategoryDescr.ColumnWidth = -2
        !SalesDptDescr.ColumnWidth = -2
        !Description.ColumnWidth = -2
        !ID.ColumnWidth = -2

        !ScanTime.ColumnOrder = 0
        !ScanDate.ColumnOrder = 0
        !CurPrice.ColumnOrder = 0
        !Quantity.ColumnOrder = 0
        !ScanCode.ColumnOrder = 0
        !CategoryDescr.ColumnOrder = 0
        !SalesDptDescr.ColumnOrder = 0
        !Description.ColumnOrder = 0
        !ID.ColumnOrder = 0
    End With

    With [Forms]![MetrologicOptimiser]![MetrologicScanBatchSubVariance].[Form]
        !txtVarianceDollars.ColumnOrder = 0
        !txtVariance.ColumnOrder = 0
        !OnShelf.ColumnOrder = 0
        !Count.ColumnOrder = 0
        !PkgSize.ColumnOrder = 0
        !Description.ColumnOrder = 0
        !ID.ColumnOrder = 0
        !CategoryDescr.ColumnOrder = 0
        !SalesDptDescr.ColumnOrder = 0

        !txtVarianceDollars.ColumnWidth = -2
        !txtVariance.ColumnWidth = -2
        !OnShelf.ColumnWidth = -2
        !Count.ColumnWidth = -2
        !PkgSize.ColumnWidth = -2
        !Description.ColumnWidth = -2
        !ID.ColumnWidth = -2
        !CategoryDescr.ColumnWidth = -2
        !SalesDptDescr.ColumnWidth = -2
        !SalesDptDescr.ColumnWidth = -2
    End With

    ShowColummnsItemsNotScanned
End Sub

Private Sub ShowColummnsItemsNotScanned()
     With [Forms]![MetrologicOptimiser]![MetrologicOptimiserSubItemsNotScanned].[Form]
        ![SalesDepartmentID].ColumnHidden = True
        !cboSalesDescription.ColumnHidden = False
        ![SubCategory.ID].ColumnHidden = True
        ![SubCategory.Description].ColumnHidden = False
        ![Item.ID].ColumnHidden = False
        ![Item.Description].ColumnHidden = False
        !Onhand.ColumnHidden = False
        !SumOfQuantity.ColumnHidden = True

        !SumOfQuantity.ColumnOrder = 0
        !Onhand.ColumnOrder = 0
        ![Item.Description].ColumnOrder = 0
        ![Item.ID].ColumnOrder = 0
        ![SubCategory.Description].ColumnOrder = 0
        ![SubCategory.ID].ColumnOrder = 0
        !cboSalesDescription.ColumnOrder = 0
        ![SalesDepartmentID].ColumnOrder = 0

        !SumOfQuantity.ColumnWidth = -2
        !Onhand.ColumnWidth = -2
        ![Item.Description].ColumnWidth = -2
        ![Item.ID].ColumnWidth = -2
        ![SubCategory.Description].ColumnWidth = -2
        ![SubCategory.ID].ColumnWidth = -2
        !cboSalesDescription.ColumnWidth = -2
        ![SalesDepartmentID].ColumnWidth = -2
    End With

End Sub
Private Sub ShowFinishButtons()

    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    With Me.MetrologicOptimiserSubAuditLog.Form
        'If gcfHandleErrors Then On Error GoTo ErrLog
        Set rs = .RecordsetClone ' Me.Recordset.Clone
        If rs.RecordCount < 1 Then
            pgAudit.Visible = False
            pgVariance.Visible = False
            pgFinish.Visible = False
        Else
            pgAudit.Visible = True
            pgVariance.Visible = True
            pgFinish.Visible = True

        End If

    End With
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub StoreTransfer()
On Error GoTo cmdStoreTransfer_Click_Err

Dim strSQL As String
    lblVendorClubAgent.Caption = "Select a Agent to create a Store Transfer"
    strSQL = "SELECT LastName as Name, AgentNumber as ID FROM Agents WHERE (((ActiveForTransfer)<>0)) ORDER BY LastName"
    With lstVendorClubAgent
        .RowSource = strSQL
        .Requery
        .Tag = "Agent"
        .BackColor = lblStoreTransfer.BackColor
        .Visible = True
    End With
    cmdContinue.Visible = True

cmdStoreTransfer_Click_Exit:
    Exit Sub

cmdStoreTransfer_Click_Err:
    MsgBox Error$
    Resume cmdStoreTransfer_Click_Exit
End Sub

Private Sub UpdateSettings()
    If gcfHandleErrors Then On Error GoTo ErrLog
    SetKMDSSettings "MetrologicSalesDepartmentBegin", cboSalesDepartmentIDBegin
    SetKMDSSettings "MetrologicSalesDepartmentEnd", cboSalesDepartmentIDEnd
    SetKMDSSettings "MetrologicCategoryBegin", cboCategoryIDBegin
    SetKMDSSettings "MetrologicCategoryEnd", cboCategoryIDEnd
    If cboCategoryIDEnd < cboCategoryIDBegin Then
        MsgBox "End Category can't be before Begin Category"
        cboCategoryIDEnd = cboCategoryIDBegin
    End If

    If cboSalesDepartmentIDEnd < cboSalesDepartmentIDBegin Then
        MsgBox "End Sales Department can't be before Begin Sales Department"
        cboSalesDepartmentIDEnd = cboSalesDepartmentIDBegin
    End If


    cboNotScannedCategoryIDBegin = cboCategoryIDBegin
    cboNotScannedCategoryIDEnd = cboCategoryIDEnd

    txtSalesDepartmentBegin = cboSalesDepartmentIDBegin
    txtSalesDepartmentEnd = cboSalesDepartmentIDEnd
    MetrologicOptimiserSubAuditLog.Requery
    MetrologicScanBatchSubVariance.Requery
    MetrologicOptimiserSubItemsNotScanned.Requery
ExitSub:
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub TabItemsNotScanned_Click()
    ShowColummnsItemsNotScanned
End Sub
