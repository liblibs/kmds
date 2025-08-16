Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbClubID_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdUseOrderGuide.Visible = Not IsNull(lstClubs)
    cmdNoOrderGuide.Visible = Not IsNull(lstClubs)

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub cboClubClone_AfterUpdate()
    Dim strTemp As String
    strTemp = "Use order guide from " & vbCrLf
    strTemp = strTemp & cboClubClone.Column(0) & " " & cboClubClone.Column(1)
    If MsgBox(strTemp, vbYesNo) = vbYes Then
        cmdCreate True, cboClubClone.Column(0)
    End If


End Sub

Private Sub cmdClone_Click()
    cboClubClone.Visible = True

End Sub

Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_Command24_Click

    DoCmd.Close

Exit_Command24_Click:
    Exit Sub

Err_Command24_Click:
    MsgBox Err.Description
    Resume Exit_Command24_Click
End Sub

Private Sub cmdEditThisOrder_Click()
    Dim nOrderID As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
If gcfHandleErrors Then On Error GoTo ErrLog
    nOrderID = lstInvoices
    stDocName = "ClubOrder"
    stLinkCriteria = "[ID]=" & nOrderID
    lbRetr.Visible = True
    Me.Repaint
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
    DoCmd.Close acForm, Me.Name
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub

End Sub

Private Sub cmdEmailOrders_Click()
    DoCmd.Close
    DoCmd.OpenForm "ClubOrderXLImport"
End Sub



Private Sub cmdNoOrderGuide_Click()
    cmdCreate False, lstClubs
End Sub

Private Sub cmdOGfromHistory_Click()
    cmdCreateFromSalesHistory

End Sub

Private Sub cmdPDADownload_Click()
Dim nOrderID As Long

    Dim stDocName As String
    Dim stLinkCriteria As String

    If MsgBox("Create Club order from PDA for " & lstClubs.Column(1), vbYesNo) = vbYes Then
        nOrderID = CreateClubOrderFromPDA(lstClubs)
        stDocName = "ClubOrder"
        stLinkCriteria = "[ID]=" & nOrderID
        DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub cmdPrevInvoices_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrevInvoices_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Club"

  '  stLinkCriteria = "[ID]=" & cmdPrevInvoices.Tag
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdPrevInvoices_Click:
    Exit Sub

Err_cmdPrevInvoices_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrevInvoices_Click
End Sub

'Private Sub cmdReprice_Click()
'    RepriceClubOrder
''Dim comQDF As New Command
''Dim varReturn
''If gcfHandleErrors Then On Error GoTo ErrLog
''
''    'added 02/01/2005
''
''
''    'This will 'reprice' all orders with vendor #1
''    'to fix the price for orders input in the previous month
''    varReturn = SysCmd(acSysCmdSetStatus, "Reprice Orders with current Pricing (qryOrderLineUpdatePrice) ")
''
''
''    With comQDF
''        .ActiveConnection = CurrentProject.Connection
''        .CommandText = "qryOrderLineUpdatePrice"
''        .CommandType = adCmdStoredProc
''        .Execute
''    End With
''    Set comQDF = Nothing
''
''
''    lstInvoices.Requery
''    lstInvoices.SetFocus
''    cmdReprice.Visible = False
''ExitSub:
''   varReturn = SysCmd(acSysCmdClearStatus)
''    Exit Sub
''
''ErrLog:
''    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
''    Resume ExitSub
'
'
'End Sub

Private Sub cmdUseOrderGuide_Click()
    cmdCreate True, lstClubs
End Sub

Private Sub cmdCreate(OG As Boolean, ClubIDForOrderGuide As Long)
    Dim nOrderID As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    lbRetr.Caption = "Creating Dispenser Order"
    lbRetr.Visible = True
    Me.Repaint

    nOrderID = CreateClubOrder(lstClubs, OG, ClubIDForOrderGuide)
    stDocName = "ClubOrder"
    stLinkCriteria = "[ID]=" & nOrderID
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
    DoCmd.Close acForm, Me.Name
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub

Private Sub cmdCreateFromSalesHistory()
    Dim nOrderID As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    lbRetr.Caption = "Creating Dispenser Order from Sales History"
    lbRetr.Visible = True
    Me.Repaint

    nOrderID = CreateClubOrderFromSalesHistory(lstClubs)
    stDocName = "ClubOrder"
    stLinkCriteria = "[ID]=" & nOrderID
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
    DoCmd.Close acForm, Me.Name
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub
End Sub



Private Sub Command45_Click()

End Sub

Private Sub cmdWebOrders_Click()
    DoCmd.Close
    DoCmd.OpenForm "ClubOrderWebImport"
End Sub

Private Sub Detail_Click()
        cmdPDADownload.Visible = True

End Sub

Private Sub Form_Click()
        cmdPDADownload.Visible = True

End Sub

Private Sub Form_DblClick(Cancel As Integer)
        cmdPDADownload.Visible = True

End Sub

Private Sub Form_Load()
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebOrders As String
Dim mstrWebDir  As String

'        Dim mstrWebDir As String
'        Dim fsotemp As New FileSystemObject

If gcfHandleErrors Then On Error GoTo ErrLog
    If gstrCashsysInsLoc = "" Then LoadDefaults

    'added 02/16/2006
    RepriceClubOrder
    lstInvoices.Requery

    cmdUseOrderGuide.Visible = False
    cmdNoOrderGuide.Visible = False
    txtSelectedClub.Visible = False

    cmdEditThisOrder.Visible = False
    txtSelectedInvoice.Visible = False
    If Me.OpenArgs = "PDA" Then
        cmdPDADownload.Visible = True
    Else
        cmdPDADownload.Visible = False
    End If



    strTemp = Dir("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\*.xls")
    If strTemp = vbNullString Then
    Else
        cmdEmailOrders.Visible = True
    End If


    strWebOrders = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
'    If strWebOrders <> "k:\agent1200\Mail" Then

        mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
        mstrWebDir = Trim(mstrWebDir)

        strTemp = mstrWebDir
        'lets build a structure for the future..
        If Not fsoTemp.FolderExists("k:\agent" & glnAgentNumber) Then
            fsoTemp.CreateFolder ("k:\agent" & glnAgentNumber)
        End If

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
'    End If

    strWebOrders = Trim(strWebOrders) & "\In\*.csv"
    strTemp = Dir(strWebOrders)
    If strTemp = vbNullString Then
        cmdWebOrders.Visible = False

    Else
        cmdWebOrders.Visible = True
    End If
ExitSub:

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


Private Sub Command24_Click()
If gcfHandleErrors Then On Error GoTo Err_Command24_Click
    DoCmd.Close

Exit_Command24_Click:
    Exit Sub

Err_Command24_Click:
    MsgBox Err.Description
    Resume Exit_Command24_Click
End Sub

Private Sub lstClubs_AfterUpdate()
Dim buildstr As String
If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedClub = lstClubs.Column(0) & " " & lstClubs.Column(1)
    cmdUseOrderGuide.Visible = Not IsNull(lstClubs)
    cmdNoOrderGuide.Visible = Not IsNull(lstClubs)
    txtSelectedClub.Visible = Not IsNull(lstClubs)
    cmdNoOrderGuide.Visible = Not IsNull(lstClubs)
    cmdOGfromHistory.Visible = Not IsNull(lstClubs)
    cmdPrintClubHistory.Visible = Not IsNull(lstClubs)
    cmdClone.Visible = True

    With cmdOrderGuide
        .Caption = "Print Guide for #" & lstClubs.Column(0)
        .Tag = lstClubs.Column(0)
        .Visible = True
    End With
'    With cmdPrevInvoices
'        .Caption = "&Previous invoices #" & lstClubs.Column(0)
'        .Tag = lstClubs.Column(0)
'        .Visible = True
'    End With
    Label49.Caption = lstClubs.Column(3)
    If Not IsNull(Trim(lstClubs.Column(2))) Or Len(Trim(lstClubs.Column(3))) > 10 Then
        buildstr = "http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q="
        buildstr = buildstr & lstClubs.Column(3)
        buildstr = buildstr & "&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
        buildstr = Replace(buildstr, " ", "+")
        lblMapQuest.Caption = "Click here for &Google Map to " & lstClubs.Column(1)
    Else
        lblMapQuest.Caption = "Click here for &Google Map"
        buildstr = "http://maps.google.com/maps"
    End If
    lblMapQuest.HyperlinkAddress = buildstr
    lblMapQuest.ControlTipText = buildstr

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub


Private Sub lstClubs_DblClick(Cancel As Integer)
    cmdCreate True, lstClubs
End Sub

Private Sub lstInvoices_AfterUpdate()
If gcfHandleErrors Then On Error GoTo ErrLog
    txtSelectedInvoice = lstInvoices.Column(1) & " " & lstInvoices.Column(3)
    cmdEditThisOrder.Visible = Not IsNull(lstInvoices)
    txtSelectedInvoice.Visible = Not IsNull(lstInvoices)
    With cmdOrderGuide
        .Caption = "Print Guide for #" & lstInvoices.Column(6)
        .Tag = lstInvoices.Column(6)
        .Visible = True
    End With
'    With cmdPrevInvoices
'        .Caption = "&Previous invoices #" & lstInvoices.Column(6)
'        .Tag = lstInvoices.Column(6)
'        .Visible = True
'    End With
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmbClubID_AfterUpdate"
    Resume ExitSub
End Sub

Private Sub lstInvoices_DblClick(Cancel As Integer)
    cmdEditThisOrder_Click
End Sub

Private Sub cmdOrderGuide_Click()
Dim stDocName As String
Dim stLinkCriteria As String

If gcfHandleErrors Then On Error GoTo Err_cmdOrderGuide_Click
    stDocName = "ClubOrderPrintDialog"
    stLinkCriteria = "[ID]=" & cmdOrderGuide.Tag
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "PrintOrderGuide"
Exit_cmdOrderGuide_Click:
    Exit Sub

Err_cmdOrderGuide_Click:
    MsgBox Err.Description
    Resume Exit_cmdOrderGuide_Click

End Sub
'------------------------------------------------------------
' cmdPrintClubHistory_Click
'
'------------------------------------------------------------
Private Sub cmdPrintClubHistory_Click()
On Error GoTo cmdPrintClubHistory_Click_Err

    DoCmd.OpenReport "Club History", acViewPreview, "", "", acNormal


cmdPrintClubHistory_Click_Exit:
    Exit Sub

cmdPrintClubHistory_Click_Err:
    MsgBox Error$
    Resume cmdPrintClubHistory_Click_Exit

End Sub
'------------------------------------------------------------
' cmdCompare_Click
'
'------------------------------------------------------------
Private Sub cmdCompare_Click()
On Error GoTo cmdCompare_Click_Err

    DoCmd.OpenForm "ClubOrderCompare", acNormal, "", "", , acDialog

cmdCompare_Click_Exit:
    Exit Sub

cmdCompare_Click_Err:
    MsgBox Error$
    Resume cmdCompare_Click_Exit

End Sub
