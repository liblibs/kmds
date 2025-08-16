Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrNow As String

Private Sub chkSendPurchaseOrder_AfterUpdate()
    lstInvoices.Visible = chkSendPurchaseOrder
    txtSelectedInvoice.Visible = chkSendPurchaseOrder
End Sub

Private Sub cmdEditThisOrder_Click()
Dim nOrderID As Long
Dim stDocName As String
Dim stLinkCriteria As String
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim nAgentNumber As Long
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim nLineCount As Integer
Dim nLineCount99 As Integer

    If gcfHandleErrors Then On Error GoTo ErrLog
    nAgentNumber = AgentNumber
    mstrNow = Format(Now, "yymmddhhnn")

    If nAgentNumber > 1000 Then nAgentNumber = nAgentNumber - 1000

    If Not fsoTemp.FolderExists("C:\NPOS") Then
        fsoTemp.CreateFolder ("C:\NPOS")
    End If

    Set txtstrTemp = fsoTemp.OpenTextFile("C:\NPOS\olf00021." & CStr(Format(nAgentNumber, "000")), ForWriting, True, TristateFalse)
    nAgentNumber = nAgentNumber + 1000

    txtstrTemp.WriteLine (Build21Header(nAgentNumber, mstrNow))
    If chkSendPurchaseOrder Then
        txtstrTemp.WriteLine (Build15Header(nAgentNumber, mstrNow))

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryCreateTransmitPurchaseOrder"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parPurchaseOrderID", adBigInt, adParamInput)
            .Parameters("parPurchaseOrderID") = CLng(lstInvoices)
            Set rstemp = .Execute
        End With

        With rstemp
            If rstemp.RecordCount = 0 Then
                MsgBox "No Records found for Order #" & lstInvoices
                GoTo ExitSub
            End If

            Do While Not .EOF
                nLineCount = nLineCount + 1
                nLineCount99 = nLineCount99 + 1
                If Not gcfHandleErrors Then Debug.Print !LineID; !ItemID
                txtstrTemp.WriteLine BuildPOLine(nAgentNumber, !ItemID, !QuantityShipped, mstrNow)
                .MoveNext
            Loop
        End With
        rstemp.Close
        Set rstemp = Nothing
        Set rstemp = CreateObject("ADODB.Recordset")

        Set comQDF = Nothing
        Set comQDF = CreateObject("ADODB.COMMAND")

        txtstrTemp.WriteLine (Build15Trailer(nAgentNumber, nLineCount, mstrNow))
    End If

    If chkSendInventory Then
        nLineCount = 0
        'create inventory file
        txtstrTemp.WriteLine (Build84Header(nAgentNumber, mstrNow))

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryCreateTransmitInventory"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            Set rstemp = .Execute
        End With

        With rstemp
            If rstemp.RecordCount = 0 Then
                MsgBox "No Records found for Sales department 1"
                GoTo ExitSub
            End If

            Do While Not .EOF
                nLineCount = nLineCount + 1
                nLineCount99 = nLineCount99 + 1
                'If not gcfHandleErrors then Debug.Print !ItemID
                txtstrTemp.WriteLine BuildInvLine(nAgentNumber, !ID, !Onhand, mstrNow)
                .MoveNext
            Loop
        End With
        rstemp.Close
        txtstrTemp.WriteLine (Build84Trailer(nAgentNumber, nLineCount, mstrNow))
    End If

    txtstrTemp.WriteLine (Build21Trailer(nAgentNumber, nLineCount99, mstrNow))
    MsgBox "Transmit File Created!" & vbCrLf & "Send the file by choosing the 'Transmit to OLCC' icon"
    DoCmd.Quit

ExitSub:

    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    DoCmd.Close acForm, Me.Name

    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub


Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    Form.InsideHeight = boxScreen.Height + 500 '+ cmdGoToFirstRecord.Height
    Form.InsideWidth = boxScreen.Width + 500
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)

End Sub

Private Sub lstInvoices_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedInvoice = lstInvoices.Column(0) & " " & lstInvoices.Column(1)

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




Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
