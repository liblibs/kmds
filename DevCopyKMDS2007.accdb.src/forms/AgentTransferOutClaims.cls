Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdEditThisOrder_Click()
    Dim nOrderID As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
If gcfHandleErrors Then On Error GoTo ErrLog
    lbRetr.Visible = True
    Me.Repaint

    nOrderID = lstInvoices
    stDocName = "AgentsTransfers"
    stLinkCriteria = "[ID]=" & nOrderID

    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, Me.Name
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")
    Dim strSQL As String
    If gbAllowDeleteDocument Then

        If MsgBox("Delete Transfer #" & lstInvoices, vbYesNo) = vbYes Then

            strSQL = "Delete from AgentTransferHead "
            strSQL = strSQL & "Where ((ID) = (" & CStr(lstInvoices) & "))"
            With comQDF
                .ActiveConnection = CurrentProject.Connection
                .CommandType = adCmdText
                .CommandText = strSQL
                .Execute
            End With
        End If
    Else
        MsgBox "This user is Not allowed to delete documents", vbCritical, glsCashierName
    End If
ExitSub:

    Set comQDF = Nothing
    lstInvoices.Requery
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

Private Sub Command4_Click()

Dim stDocName As String
Dim nOrderID As Integer
    Dim stLinkCriteria As String

If gcfHandleErrors Then On Error GoTo ErrLog
    lbRetr.Caption = "Building New Transfer..."
    lbRetr.Visible = True
    Me.Repaint


    nOrderID = CreateTransfer(lstVendors, "C")
    If nOrderID > 0 Then
        stDocName = "AgentsTransfers"
        stLinkCriteria = "[ID]=" & nOrderID
        DoCmd.OpenForm stDocName, , , stLinkCriteria
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


Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    If lstVendors.ListCount = 0 Then
        SetOptions False
        DoCmd.OpenQuery ("qryAgentsAppendOLCCClaims")
        SetOptions True
        lstVendors.Requery
    End If
    If lstVendors.ListCount = 0 Then
        MsgBox "You must set up an agent named 'OLCC CLAIMS' in the Agent screen (menu option E-6)"
        'Cancel = True
    End If
    SQLPassThru ("sp_Upd_TransferLinesWithCurrentPrice")

End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
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
    cmdfrmDeleteRecord.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub lstInvoices_DblClick(Cancel As Integer)
cmdEditThisOrder_Click
End Sub

Private Sub lstVendors_AfterUpdate()

If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedVendor = lstVendors.Column(0) & " " & lstVendors.Column(1)
    txtSelectedVendor.Visible = Not IsNull(lstVendors)
    Command4.Visible = Not IsNull(lstVendors)
    'optBasePO.Visible = Not IsNull(lstVendors)
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "lstVendors_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub lstVendors_DblClick(Cancel As Integer)
Command4_Click
End Sub

Private Sub optBasePO_Click()
If gcfHandleErrors Then On Error GoTo ErrLog
    If optBasePO = 2 Then
        dtpDateRangeTo.Visible = True
        dtpDateRangeFrom.Visible = True
        lblDateRangeTo.Visible = True
        lblDateRangeFrom.Visible = True
        lblDateRange.Visible = True
        boxDateRange.Visible = True
        dtpDateRangeFrom = CDate(DateAdd("d", -7, Now))
        dtpDateRangeTo = Now

    Else
        dtpDateRangeTo.Visible = False
        dtpDateRangeFrom.Visible = False
        lblDateRangeTo.Visible = False
        lblDateRangeFrom.Visible = False
        lblDateRange.Visible = False
        boxDateRange.Visible = False
    End If
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

Private Sub optSortBy_AfterUpdate()
    lstVendors.RowSource = "SELECT Agents.AgentNumber, Agents.LastName AS Name, Agents.FirstName, Agents.ActiveForTransfer FROM Agents WHERE (((Agents.ActiveForTransfer)<>0)) "

    If optSortBy Then
        lstVendors.RowSource = lstVendors.RowSource & " ORDER BY Agents.LastName "
        lblSortby.Caption = "Sort by Number"
    Else
        lstVendors.RowSource = lstVendors.RowSource & " ORDER BY Agents.AgentNumber "
        lblSortby.Caption = "Sort by Name"
    End If
    lstVendors.Requery
End Sub
