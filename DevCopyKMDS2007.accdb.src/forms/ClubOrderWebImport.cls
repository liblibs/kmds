Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim fsoTemp As New FileSystemObject
Dim mstrWebDir As String
Dim mnInterval As Integer
Dim mnPollCount As Integer
Dim mstrOrderNotes As String



Private Sub cmbLicensee_AfterUpdate()
    ClubOrderGuidesWeb.Requery

End Sub

Private Sub cmdCheckNow_Click()
    'force the timer routine to check..
    mnInterval = 1
End Sub

Private Sub cmdCreateTextFiles_Click()
    ExportFilesForWeb
    MsgBox "Files Created" & vbCrLf & "They will be sent during next Poll."
End Sub

Private Sub cmdOGExport_Click()
    If IsNull(cmbLicensee) Then
        MsgBox ("First Select a club to refresh")
    Else
        If (MsgBox("Replace Online Order Guide for " & vbCrLf & cmbLicensee.Column(1) & " ?", vbYesNo) = vbYes) Then
            ExportOrderGuideFileForWeb
            MsgBox ("Order Guide created!" & vbCrLf & vbCrLf & "It will be sent during next Poll." & vbCrLf & "Only 1 order guide per POLL is refreshed at the website.")
        End If
    End If

End Sub

Private Sub cmdPause_Click()
    'defeat the timer function by setting timerinterval = 0
    If cmdPause.Caption = "Pause" Then
        Me.TimerInterval = 0
        txtInterval.BackColor = vbRed
        cmdPause.Caption = "Resume...."
    Else
        '1000 ms = 1 second
        txtInterval.BackColor = txtFileName.BackColor
        Me.TimerInterval = 1000
        cmdPause.Caption = "Pause"
    End If
End Sub

Private Sub Form_Load()
    Me.InsideHeight = boxScreen.Height + BoxScreenTop.Height + 500
    Me.InsideWidth = boxScreen.Width + 500

End Sub

Private Sub Form_Open(Cancel As Integer)
    If GetKMDSSettings("IOWSub", "False") <> "True" Then
        Cancel = True
    Else
        BuildWebMailDirectoryStructure
        mnInterval = 50 ' seconds
        txtIntervalSet = 5  'minutes
        Me.TimerInterval = 1000   ' 1 second
    End If

End Sub


Private Sub Form_Resize()

'ClubOrderWebImportSub.Width = Me.InsideWidth - 300
boxScreen.Width = Me.InsideWidth - 100
BoxScreenTop.Width = boxScreen.Width
tabStatusForms.Width = boxScreen.Width - 200
WebOrderStatus.Width = boxScreen.Width - 450

End Sub

Private Sub Form_Timer()
    txtTime = Now()
    mnInterval = mnInterval - 1
    txtInterval = mnInterval
    If mnInterval <= 0 Then
        Me.TimerInterval = 0 'turn off the timer during this processing
        CheckForOrders
        mnInterval = txtIntervalSet * 60 'reset the mnInterval counter in seconds
        Me.TimerInterval = 1000
    End If
End Sub



Private Sub CheckForOrders()
    Dim strTemp As String
    Dim strWebOrders As String
    Dim strPassString As String
    mnPollCount = mnPollCount + 1
    strTemp = mstrWebDir
    strWebOrders = Trim(mstrWebDir) & "\In\*.csv"
    strTemp = Dir(strWebOrders)
    strPassString = "c:\npos\ftpweb.bat "
    WaitForExecCmd strPassString
    Do While strTemp <> vbNullString
        'Text0 = strTemp & vbcrlf & Text0
        strTemp = Dir
    Loop
    ProcessListOfCSVFiles
    txtPollCount = mnPollCount
    'Text0 = ""
End Sub

Sub ProcessListOfCSVFiles()
Dim strTemp As String
Dim strWebOrders As String
Dim strDestError As String
'   FYI... mstrWebDir = "k:\agent1200\Mail"
    strWebOrders = mstrWebDir & "\In\*.csv"
    strTemp = Dir(strWebOrders)
    Do While strTemp <> vbNullString
        txtFileName = (mstrWebDir & "\In\" & strTemp)
        ImportClubOrderFromCSV txtFileName
        If fsoTemp.FileExists(txtFileName) Then  'IT DIDN'T GET MOVED... MOVE IT TO ERROR FILE
            strDestError = Replace(txtFileName, "\IN\", "\IN\ERROR\" & Format(Now, "yyyymmdd hh nn ss"))
            fsoTemp.MoveFile txtFileName, strDestError
            WebLogEntry -999, ClubName, "ERROR PROCESSING : = " & strDestError, "-99", ""
        End If
        strTemp = Dir
        WebOrderStatus.Requery
    Loop
End Sub


Private Sub ImportClubOrderFromCSV(CSVFile As String)

Dim gstDocName As String
Dim gstLinkCriteria As String
Dim gstrInvoiceNo As String
Dim nInvoiceNo As Long
Dim nClubID As Long
Dim rs As Object
Dim rSOrdhead As Object
Dim strSQL As String
Dim comQDF As Object
Dim nOrderNumber As Long
Dim strDest As String
Dim strReturned As String
Dim strTXTFileSource As String
Dim strTXTFileDest As String
Dim nQuantity As Integer
Dim txtstrBody As TextStream
Dim coleMail As New Collection
Dim strTextBody As String

    Set comQDF = CreateObject("ADODB.COMMAND")

    strDest = Replace(CSVFile, "\IN\", "\IN\ARCHIVE\" & Format(Now, "yyyymmdd hh nn ss"))
    'where the notes are kept
    strTXTFileSource = Replace(CSVFile, ".CSV", ".TXT")
    strTXTFileDest = Replace(strDest, ".CSV", ".TXT")
    If gcfHandleErrors Then On Error GoTo ErrLog
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Delete * from ClubOrderCSV"
        .CommandType = adCmdText
        .Execute
    End With
    Set comQDF = Nothing

    DoCmd.TransferText acImportDelim, "ClubOrderCSVImportSpecification", "ClubOrderCSV", CSVFile, True
    Me.Requery ' to get the club id
    Me.Refresh
    'CREATE THE ORDER HERE!
    nOrderNumber = CreateClubOrder(ClubID, False, ClubID)
    txtOrderNumber = nOrderNumber

    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderCSVAppend"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parOrderNumber", adBigInt, adParamInput)
        .Parameters("parOrderNumber") = nOrderNumber
        .Execute
    End With
    Set comQDF = Nothing
    fsoTemp.MoveFile CSVFile, strDest

    Set rs = CreateObject("ADODB.Recordset")
    Set rSOrdhead = CreateObject("ADODB.Recordset")
    If gcfHandleErrors Then On Error GoTo ErrLog
    ' get the totals and update the ordhead...
    Set rs = CurrentDb().OpenRecordset("select * from  qryClubOrderLineTotals where OrdheadID =" & nOrderNumber, dbReadOnly)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        strSQL = "SELECT * "
        strSQL = strSQL & "FROM Ordhead "
        strSQL = strSQL & "WHERE (Id =" & nOrderNumber & ")"

        rSOrdhead.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rSOrdhead.BOF And rSOrdhead.EOF Then
            MsgBox "There is no Order found for : " & rSOrdhead
            Exit Sub
        End If
        With rSOrdhead
            !InvDollars = rs!ExtPrice               'Gross Invoice
            !DisDollars = rs!ExtDiscount            'Gross Discount
            !OrdDollars = rs!ExtNetPrice            ' net amount of invoice... amt customer is paying..
            !InvQty = rs!Quantity                   ' total qty
            !LiquorDollars = rs!LiquorExtPrice      'Gross Liquor
            !LiquorDiscount = rs!LiquorDiscount     'Discount Liquor
            !NonLiquorDollars = rs!NonLiquorExtPrice   'Gross Related
            !NonLiquorDiscount = rs!NonLiquorDiscount  'Discount Related
            nInvoiceNo = Trim(!InvoiceNo)
            nClubID = !ClubID
            nQuantity = !InvQty
            'add the notes

            .Update
        End With
        rSOrdhead.Close
        Set rSOrdhead = Nothing
    End If
    rs.Close
    Set rs = Nothing

'save the notes here so if there is an error the order has been saved... just not printed..
    'Open the file for reading

    Set txtstrBody = fsoTemp.OpenTextFile(strTXTFileSource, ForReading)
    'The ReadAll method reads the entire file into the variable BodyText
    With txtstrBody
        Dim strLicenseeEmail As String

        strLicenseeEmail = .ReadLine
        .ReadLine ' order number in website's Ordhead
        mstrOrderNotes = .ReadAll
        mstrOrderNotes = Replace(mstrOrderNotes, "Type in any notes you would like to send to the store and then click on Submit to send Order.", "-")

        'Close the file
        .Close
    End With
    Set txtstrBody = Nothing
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM Ordhead "
    strSQL = strSQL & "WHERE (Id =" & nOrderNumber & ")"
    Set rSOrdhead = CreateObject("ADODB.Recordset")

    rSOrdhead.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rSOrdhead.BOF And rSOrdhead.EOF Then
        MsgBox "There is no Order found for : " & rSOrdhead
        Exit Sub
    End If
    With rSOrdhead
        !OrderNotes = mstrOrderNotes
        .Update
    End With
    rSOrdhead.Close
    Set rSOrdhead = Nothing

    fsoTemp.MoveFile strTXTFileSource, strTXTFileDest

    Me.Refresh
    If Not gcfHandleErrors Then Debug.Print "Invoice = " & nInvoiceNo
       'cmbOrderID.Requery
       'Me.Refresh
    gstDocName = "ClubOrderPrintDialog"
    gstLinkCriteria = "[ID]=" & nClubID
    gstrInvoiceNo = nInvoiceNo
   ' MsgBox "ClubOrderPrintDialog"
    DoCmd.OpenForm gstDocName, , , gstLinkCriteria, , , gstrInvoiceNo
    Dim stDocName As String
    stDocName = "ClubInvoiceStoreCopy"
    ArchiveReportDocument stDocName, stDocName, CStr(nClubID), CStr(nInvoiceNo)
    ArchiveReport stDocName, CStr(nInvoiceNo) & " " & CStr(nInvoiceNo)
    DoCmd.OpenReport stDocName, acViewPreview
    DoCmd.PrintOut , , , , 1 'txtStoreInvoices
    DoCmd.Close acReport, stDocName
    DoCmd.Close acForm, gstDocName

    strTextBody = "We have received your order of "
    '& strBottledesc & vbcrlf
    If nQuantity = 1 Then
        strTextBody = strTextBody & "1 item." & vbCrLf
    Else
        strTextBody = strTextBody & CStr(nQuantity) & " items." & vbCrLf
    End If

    strTextBody = strTextBody & "It is entered into our system with Invoice #" & CStr(nInvoiceNo) & "." & vbCrLf
    strTextBody = strTextBody & "We will contact you with a final total once we have it ready for pick up." & vbCrLf
    strTextBody = strTextBody & "Thank you for your business." & vbCrLf

    KMDSeMail strLicenseeEmail, coleMail, "Web Order Confirmation Inv#" & CStr(nInvoiceNo), strTextBody


    WebLogEntry nInvoiceNo, ClubName, mstrOrderNotes, CStr(nQuantity), strLicenseeEmail


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub


Sub ExportOrderGuideFileForWeb()

Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String
Dim nClubID As Long


    If gcfHandleErrors Then On Error GoTo ErrLog


    nClubID = CLng(cmbLicensee)
    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\OrderGuide.txt" ' k:\agent1200\Mail\Out\Item.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")


    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderGuideWebExport"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("parClubID", adBigInt, adParamInput)
        '.Parameters.Append .CreateParameter("parCashierID", adBigInt, adParamInput)

        .Parameters("parClubID") = nClubID

        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        rstemp.MoveFirst

    End With


    With rstemp
        Do While Not .EOF
            If ((!ClubID = nClubID) And (!OrderID = 1)) Then
                strDesc = IIf(IsNull(!Description), "", !Description)
                strDesc = Replace(strDesc, "'", "")
                strDesc = Replace(strDesc, Chr(34), "")
                txtstrTemp.WriteLine !ClubID
                txtstrTemp.WriteLine !OrderID
                txtstrTemp.WriteLine !LiquorItem
                txtstrTemp.WriteLine !ItemID
                txtstrTemp.WriteLine strDesc
                txtstrTemp.WriteLine !Par
                txtstrTemp.WriteLine !QuotedPrice
            End If
            .MoveNext
        Loop
    End With
    rstemp.Close
ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportOrderGuideFileForWeb"
    Resume ExitSub
End Sub

Sub BuildWebMailDirectoryStructure()
Dim strTemp As String
Dim strWebOrders As String

 On Error GoTo ErrLog

    If Not fsoTemp.FolderExists("k:\agent" & glnAgentNumber) Then
        fsoTemp.CreateFolder ("k:\agent" & glnAgentNumber)
    End If


    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)

    strTemp = mstrWebDir
    'build a structure for the future..
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

    If Not fsoTemp.FolderExists(strTemp & "\Out") Then
        fsoTemp.CreateFolder (strTemp & "\Out")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\Out\Archive") Then
        fsoTemp.CreateFolder (strTemp & "\Out\Archive")
    End If
    If Not fsoTemp.FolderExists(strTemp & "\Out\ERROR") Then
        fsoTemp.CreateFolder (strTemp & "\Out\ERROR")
    End If
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "BuildWebMailDirectoryStructure"
    Resume ExitSub

End Sub





Private Sub WebLogEntry(WebOrderID As Long, LicenseeName As String, WebOrderNotes As String, BottleCount As String, CustomerEmail As String)


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


'Dim adoCnOLCNetwork As New ADODB.Connection
Dim adoCnOLCNetwork As Object
 On Error GoTo ErrLog

Set adoCnOLCNetwork = CreateObject("ADODB.Connection")



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

    With comQDF
        .CommandTimeout = 15
        .ActiveConnection = adoCnOLCNetwork
        .CommandText = "sp_WebLog"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular


        .Parameters.Append .CreateParameter("parSubroutine", adChar, adParamInput, 50)
        .Parameters.Append .CreateParameter("parSource", adChar, adParamInput, 50)
        .Parameters.Append .CreateParameter("parErrorCode", adChar, adParamInput, 50)
        .Parameters.Append .CreateParameter("parLicenseeName", adChar, adParamInput, 50)
        .Parameters.Append .CreateParameter("parWebOrderNotes", adChar, adParamInput, 250)
        .Parameters.Append .CreateParameter("parWebOrderID", adInteger, adParamInput)

        .Parameters("parSubroutine") = CustomerEmail
        .Parameters("parSource") = BottleCount
        .Parameters("parErrorCode") = "-1"
        .Parameters("parLicenseeName") = LicenseeName
        .Parameters("parWebOrderNotes") = WebOrderNotes
        .Parameters("parWebOrderID") = WebOrderID
        .Execute
    End With
    Set comQDF = Nothing
    adoCnOLCNetwork.Close
    Set adoCnOLCNetwork = Nothing
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "WebLogEntry"
    Resume ExitSub
End Sub
