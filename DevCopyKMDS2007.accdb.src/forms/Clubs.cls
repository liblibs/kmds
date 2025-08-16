Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub cboID_AfterUpdate()
   ' Find the record that matches the control.
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Me![cboID])
    Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "cboID_AfterUpdate"
    Resume ExitSub

End Sub

Private Sub cmdNoOrderGuide_Click()
    cmdCreate False, ID
End Sub

Private Sub cmdUseOrderGuide_Click()
    cmdCreate True, ID
End Sub

Private Sub cmdCreate(OG As Boolean, ClubIDForOrderGuide As Long)
    Dim nOrderID As Long
    Dim stDocName As String
    Dim stLinkCriteria As String
If gcfHandleErrors Then On Error GoTo ErrLog
    nOrderID = CreateClubOrder(ID, OG, ClubIDForOrderGuide)
    stDocName = "ClubOrder"
    stLinkCriteria = "[ID]=" & nOrderID
    DoCmd.Close
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "ClubOrderCreate"
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdCreate"
    Resume ExitSub

End Sub

Private Sub cmbLicensee_AfterUpdate()
    If Not IsNull(cmbLicensee) Then
        If MsgBox("Use this Licensee Info ?", vbYesNo) = vbYes Then
            Licensee = cmbLicensee
            If optLicenseeSort.Value = 1 Then '1 = sort by name so name is in column 0
                txtName = cmbLicensee.Column(0)
            Else
                txtName = cmbLicensee.Column(1)
            End If
            Address1 = cmbLicensee.Column(2)
            City = cmbLicensee.Column(3)
            St = "OR"
            Zip = cmbLicensee.Column(4)
            Form_Dirty False
        End If
    End If
End Sub

Private Sub cmbLicensee_Click()
'    If Me.Dirty Then
'        Me.Dirty = False
'    End If
End Sub

Private Sub cmbLicensee_GotFocus()
'    If Me.Dirty Then
'        Me.Dirty = False
'    End If

End Sub

Private Sub cmdfrmAddNew_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    MsgBox "Adding a new club.. Find the club in the Licensee list"
    DoCmd.GoToRecord , , acNewRec
    cmbLicensee.SetFocus


ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmAddNew_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo ErrLog


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    'set focus to combo
    cboID.Visible = True
    cboID.SetFocus

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmDeleteRecord_Click"
    Resume ExitSub

End Sub

Private Sub cmdfrmSaveRecord_Click()
Dim strTemp As String
If gcfHandleErrors Then On Error GoTo ErrLog

    'strTemp = Desc
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    'refresh combo box,
    cboID.Requery
    cboID = ID
    cboID.Visible = True
'    cboID = strTemp
'    cboID.SetFocus
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmSaveRecord_Click"
    Resume ExitSub
End Sub

Private Sub cmdfrmUndo_Click()
Dim ctrlTemp As Control

If gcfHandleErrors Then On Error GoTo ErrLog

    Set ctrlTemp = Screen.PreviousControl

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    ctrlTemp.SetFocus
    'cmdfrmUndo.Visible = False
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmUndo_Click"
    Resume ExitSub

End Sub

Private Sub cmdPalmFrom_Click()

   Dim rstemp As Recordset


    'Dim rsOrder As New Command
    Dim rsOrder As Object
    Set rsOrder = CreateObject("ADODB.Recordset")

    Dim strSQL As String
    Dim nLineID As Long
    Dim strCnn As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    If IsNull(gstrPalmDataDirectoryInventoryFrom) Or gstrPalmDataDirectoryInventoryFrom = "" Then
        MsgBox "Palm Data Directory Inventory From  not set in system file (G-5)"
        Exit Sub
    End If
    If Dir(gstrPalmDataDirectoryInventoryFrom) = "" Then
        MsgBox gstrPalmDataDirectoryInventoryFrom & " does not exist!"
        Exit Sub
    End If


    strSQL = "This will save the order guide in the order it is currently displayed." & vbCrLf & vbCrLf
    strSQL = strSQL & "Do you want to save the order guide now?"

    If MsgBox(strSQL, vbYesNo) = vbNo Then
        Exit Sub
    End If

    If gcfHandleErrors Then On Error GoTo Err_cmdPalmFrom_Click

    Dim stDocName As String

    SetOptions False
    stDocName = "qryClubOrderGuideFromPalm_Step_1"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
    If gbPalmCommaDelimited Then
        stDocName = "PalmFrom Import Specification Comma"
    Else
        stDocName = "PalmFrom Import Specification Tab" 'there is an extra spacen in this one..
    End If

    DoCmd.TransferText acImportDelim, stDocName, "PalmOrderGuideFrom", gstrPalmDataDirectoryInventoryFrom, False

    stDocName = "qryClubOrderGuideFromPalm_Step_2"
    DoCmd.OpenQuery stDocName, acNormal, acEdit


    Kill gstrPalmDataDirectoryInventoryFrom


    '1. open the grid as a recordset
    Set rstemp = CurrentDb().OpenRecordset("qryClubOrderGuideFromPalm_Step_3")

    'Set rs = CurrentDb().OpenRecordset(strSQL)

    '2. save the grid as a -1 order guide records for this club
    '   Note: an error in this routine indicates the previous save did NOT complete
    If gcfHandleErrors Then On Error GoTo ErrPreviousSave
    With rstemp
        If .RecordCount > 0 Then
            .MoveFirst
            rsOrder.Open "ClubOrderGuides", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            Do While Not .EOF
                nLineID = nLineID + 1
                rsOrder.AddNew
                rsOrder!ClubID = ID
                rsOrder!OrderID = -1
                rsOrder!LineID = nLineID
                rsOrder!LiquorItem = !Liquor ' 'True
                rsOrder!ItemID = !ItemID
                'rsOrder!Description = " "
                'rsOrder!Par = !Physical  'stuffed in here when creating order Par
                rsOrder!QuotedPrice = 0 '!QuotedPrice
                ' is this necessary? rsOrder!Description = !Description
                '233
                rsOrder.Update
                If Not gcfHandleErrors Then Debug.Print .AbsolutePosition + 1;
                .MoveNext
            Loop
            rsOrder.Close
            .Close
        End If
    End With
    Set rstemp = Nothing
    Set rsOrder = Nothing

    If gcfHandleErrors Then On Error GoTo ErrRollBack
  '  cnn1.BeginTrans

    '3. Delete original order guide records for this club
    strSQL = "DELETE  "
    strSQL = strSQL & "FROM ClubOrderGuides "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)=" & ID & ") "
    strSQL = strSQL & " AND ((ClubOrderGuides.OrderID)=1))"
    CurrentDb().Execute strSQL
   ' With comQDF
   '     .ActiveConnection = cnn1
   '     .CommandText = strSQL
   '     .CommandType = adCmdText
    '    .Execute , , adCmdText
    'End With


    '4. change the -1 order guide records to 1 order guide
    strSQL = "UPDATE ClubOrderGuides "
    strSQL = strSQL & "SET ClubOrderGuides.OrderID = 1 "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)=" & ID & ") "
    strSQL = strSQL & "AND ((ClubOrderGuides.OrderID)=-1))"
    CurrentDb().Execute strSQL
    'With comQDF1
    '    .ActiveConnection = cnn1
    '    .CommandText = strSQL
    '    .CommandType = adCmdText
    '    .Execute , , adCmdText
    'End With

    '5. Delete any -1 order guide records for this club
    strSQL = "DELETE  "
    strSQL = strSQL & "FROM ClubOrderGuides "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)=" & ID & ") "
    strSQL = strSQL & " AND ((ClubOrderGuides.OrderID)=-1))"
    CurrentDb().Execute strSQL

    'With comQDF2
    '    .ActiveConnection = cnn1
    '    .CommandText = strSQL
    '    .CommandType = adCmdText
    '    .Execute , , adCmdText
    'End With
    'cnn1.CommitTrans
    'cnn1.Close
    'this forces the subform to resort

    'optOrderBy = 1
    'optOrderBy_Click

ExitSub:
    'Set comQDF = Nothing
    'Set comQDF1 = Nothing
    'Set comQDF2 = Nothing
    'Set cnn1 = Nothing
    Exit Sub

ErrRollBack:
    strSQL = "There were errors saving the order guide for Club: " & ID
    MsgBox strSQL
    'cnn1.RollbackTrans
    'cnn1.Close
    Resume ExitSub

ErrLog:
    EMsgBox Me, "cmdSaveOrderGuide_Click"
    Resume ExitSub

ErrPreviousSave:
    strSQL = "There were errors saving the order guide for Club: " & ID
    MsgBox strSQL
    EMsgBox Me, "CommcmdSaveOrderGuide_Clickand116"

    Resume ExitSub








    Me.Requery
    SetOptions True

Exit_cmdPalmFrom_Click:
    Exit Sub

Err_cmdPalmFrom_Click:
    MsgBox Err.Description
    Resume Exit_cmdPalmFrom_Click

End Sub


Private Sub cmdRefreshLicensee_Click()
    SetOptions False
        'added 31415
    If (GetKMDSSettings("UseWinSCPForOLASTransmits", "False")) = True Then
        'build a Dayend files
        '1. WinSCPDaily.txt 'script file used by :
        '2. WinSCPDaily.bat
        CreateWinSCPtoOLASFTPScriptFile ("Licensee")
        WaitForExecCmd "c:\npos\WinSCPLicensee.BAT"
    Else
        OLCCFTPGet "licensee.txt", "\\" & gstrCashsysInsLoc & "\npos\licensee.txt"
    End If

    DoCmd.RunSQL "Delete * From LocLicensee", False
    DoCmd.TransferText acImportFixed, "LICENSEE Import Specification", "LocLicensee", "\\" & gstrCashsysInsLoc & "\npos\licensee.txt", False
    DoCmd.OpenQuery "qryLicensee_01_SetPremisesNumber"
    DoCmd.OpenQuery "qryLicensee_02_AppendFromLoc"
    DoCmd.OpenQuery "qryLicensee_03_UpdateFromLoc"
    MsgBox "Refreshed!"
    SetOptions True
End Sub

Private Sub Desc_BeforeUpdate(Cancel As Integer)
    'since club/vendor names are used in the file structure
    'they can not have illegal file names!Dim ni As Integer
    Dim nI As Integer
    For nI = 1 To Len(txtName)
        If InStr(gstrIllegalFname, Mid(txtName, nI, 1)) Then
            MsgBox "this name can not contain any of these characters " & vbCrLf & gstrIllegalFname
            Cancel = True
            Exit Sub
        End If
    Next nI
End Sub

Private Sub Detail_DblClick(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
'    With cdlgClubInvoices
'        .InitDir = gstrReportsDir
'        If gbSaveReportsAsPDF Then
'            .Filter = "Adobe Acrobat (*.pdf)|*.pdf"
'        Else
'            .Filter = "Snapshots (*.snp)|*.snp"
'        End If
'        .ShowOpen
'        MsgBox .filename
'    End With
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Detail_DblClick"
    Resume ExitSub

End Sub

Private Sub Form_AfterInsert()
    cboID.Requery
    Me.Requery
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

    Dim con As Object
    Dim rs As Object
    Dim stSql As String
    Dim lngClubID As Long

If gcfHandleErrors Then On Error GoTo ErrLog
    If IsNull(Licensee) Then
        Licensee = 0
    End If
    If Not IsNumeric(Licensee) Then
        Licensee = 0
    End If


'    ' Assign the next available club id
'
'    Set con = Application.CurrentProject.Connection
'    Set rs = CreateObject("ADODB.Recordset")
'    stSql = "SELECT Max(Club.ID) AS MaxOfID FROM Club "
'    rs.Open stSql, con, 1    ' 1 = adOpenKeyset
'
'    ' If no item matches, report the error and exit the function.
'    If (rs.EOF) Then
'        lngClubID = 100
'    Else
'        rs.MoveFirst
'        lngClubID = rs!maxofid
'    End If
'    'ID = lngClubID + 1

    ChangeDate = Now
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_BeforeUpdate"
    Resume ExitSub
End Sub

Private Sub Form_Current()
Dim rs As Recordset
Dim strSQL As String
Dim buildstr As String
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdMailOrderGuide.Visible = False
    If Me.NewRecord Then
        ID.Visible = True
        ID.SetFocus
        cboID.Visible = False
        strSQL = "SELECT Max(Club.ID) AS MaxOfID "
        strSQL = strSQL & "FROM Club" ' WHERE (((Club.ID)<2000))"

        Set rs = CurrentDb().OpenRecordset(strSQL)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            ID = rs!maxofid + 1
            rs.Close
        Else
            ID = 100
        End If
        Set rs = Nothing
        cmbLicensee = "" 'Licensee
        ID.Locked = False


    Else
        cboID.Visible = True
        cboID.SetFocus

        'ID.Visible = False
        cboID = ID
        ID.Locked = True
        cmbLicensee = Licensee
        If Not IsNull(EMail) And EMail <> "" Then
            cmdMailOrderGuide.Visible = True
        End If
'        'this works for map quest
'        buildstr = "http://www.mapquest.com/maps/map.adp?formtype="
'        buildstr = buildstr & "&address=" & Address1
'        buildstr = buildstr & "&city=" & City
'        buildstr = buildstr & "&state=" & St
'        buildstr = buildstr & "&zipcode=" & Zip
'        lblMapQuest.HyperlinkAddress = buildstr
'        lblMapQuest.Caption = "Click here for &MapQuest to " & Desc

        'http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=32994+HWY+99E,+TANGENT,+OR+97389&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1
        If Not IsNull(Trim(Address1)) Or Trim(Address1) <> "" Then
            buildstr = "http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q="
            buildstr = buildstr & Trim(Address1) & ", " & Trim(City) & ", " & Trim(St) & " " & Trim(Zip)
            buildstr = buildstr & "&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
            buildstr = Replace(buildstr, " ", "+")
            lblMapQuest.Caption = "Click here for &Google Map to " & txtName
        Else
            lblMapQuest.Caption = "Click here for &Google Map"
            buildstr = "http://maps.google.com/maps"
        End If
        If Not gcfHandleErrors Then Debug.Print buildstr
        lblMapQuest.HyperlinkAddress = buildstr
        lblMapQuest.ControlTipText = buildstr


    End If

    cmdfrmAddNew.Visible = Not Me.NewRecord
    cmdfrmDeleteRecord.Visible = Not Me.NewRecord
    cmdfrmSaveRecord.Visible = False
    cmdfrmUndo.Visible = False

    If Not Me.NewRecord And optObsolete = -1 Then ' obsolete
        txtName.BackColor = vbRed
        Licensee.BackColor = vbRed
        With lblObsolete
            .ForeColor = vbRed
            .Caption = "This Licensee is OBSOLETE and will not appear in Create Order List"
        End With
        Box135.Visible = True

    Else
        txtName.BackColor = Address1.BackColor
        Licensee.BackColor = Address1.BackColor
        With lblObsolete
            .ForeColor = Label72.ForeColor
            .Caption = "Set this Licensee to OBSOLETE"
        End With
        Box135.Visible = False
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub

End Sub

Private Sub Form_Delete(Cancel As Integer)
    'does this club have an order guide?
Dim strSQL As String
Dim rs As Recordset
    strSQL = "SELECT ClubOrderGuides.ClubID "
    strSQL = strSQL & "FROM ClubOrderGuides "
    strSQL = strSQL & "WHERE (((ClubOrderGuides.ClubID)= " & ID & "))"
    Set rs = CurrentDb().OpenRecordset(strSQL)
    If rs.RecordCount = 0 Then
        If MsgBox("Are you sure?", vbYesNo, "Delete This Record") = vbNo Then
            Cancel = True
        End If
    Else
        MsgBox ("You need to delete the order guide for this club before you can delelte this club")
        MsgBox ("Do this by saving the order guide with a blank order in the Create Club Order Screen")
        Cancel = True
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Dirty(Cancel As Integer)
If gcfHandleErrors Then On Error GoTo ErrLog
    cmdfrmSaveRecord.Visible = True
    cmdfrmUndo.Visible = True
ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Dirty"
    Resume ExitSub

End Sub



Private Sub cmdGoToNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToNew_Click

    DoCmd.GoToRecord , , acNewRec
    txtName.SetFocus
Exit_cmdGoToNew_Click:
    Exit Sub

Err_cmdGoToNew_Click:
    EMsgBox Me, "cmdGoToNew_Click"
    Resume Exit_cmdGoToNew_Click

End Sub

Private Sub cmdDeleteRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdDeleteRecord_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
    txtName.SetFocus

Exit_cmdDeleteRecord_Click:
    Exit Sub

Err_cmdDeleteRecord_Click:
    Select Case Err.Number
    Case 2501
    Case Else
        EMsgBox Me, "cmdDeleteRecord_Click"
    End Select
    Resume Exit_cmdDeleteRecord_Click

End Sub

Private Sub cmdGoToLastRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToLastRecord_Click


    DoCmd.GoToRecord , , acLast
    txtName.SetFocus

Exit_cmdGoToLastRecord_Click:
    Exit Sub

Err_cmdGoToLastRecord_Click:
    EMsgBox Me, "cmdGoToLastRecord_Click"
    Resume Exit_cmdGoToLastRecord_Click

End Sub
Private Sub cmdGoToNext_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToNext_Click


    DoCmd.GoToRecord , , acNext
    txtName.SetFocus
Exit_cmdGoToNext_Click:
    Exit Sub

Err_cmdGoToNext_Click:
    Select Case Err.Number
    Case 2105
        Beep
    Case Else
        EMsgBox Me, "cmdGoToNext_Click"
    End Select
    Resume Exit_cmdGoToNext_Click

End Sub
Private Sub cmdGoToPreviousRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToPreviousRecord_Click


    DoCmd.GoToRecord , , acPrevious
    txtName.SetFocus

Exit_cmdGoToPreviousRecord_Click:
    Exit Sub

Err_cmdGoToPreviousRecord_Click:
    Select Case Err.Number
    Case 2105
        Beep
    Case Else
        EMsgBox Me, "cmdGoToPreviousRecord_Click"
    End Select
    Resume Exit_cmdGoToPreviousRecord_Click

End Sub

Private Sub cmdGoToFirstRecord_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdGoToFirstRecord_Click


    DoCmd.GoToRecord , , acFirst
    txtName.SetFocus

Exit_cmdGoToFirstRecord_Click:
    Exit Sub

Err_cmdGoToFirstRecord_Click:
    EMsgBox Me, "cmdGoToFirstRecord_Click"
    Resume Exit_cmdGoToFirstRecord_Click

End Sub
'Private Sub cmdFRMSaveRecord_Click()
'If gcfHandleErrors Then On Error GoTo Err_cmdFRMSaveRecord_Click
'
'
'    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
'    cboID.Requery
'    cboID = ID
'    cboID.Visible = True
'    cboID.SetFocus
'
'Exit_cmdFRMSaveRecord_Click:
'    Exit Sub
'
'Err_cmdFRMSaveRecord_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdFRMSaveRecord_Click
'
'End Sub
'Private Sub cmdFRMUndo_Click()
'If gcfHandleErrors Then On Error GoTo Err_cmdFRMUndo_Click
'
'
'    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
'
'Exit_cmdFRMUndo_Click:
'    Exit Sub
'
'Err_cmdFRMUndo_Click:
'    MsgBox Err.Description
'    Resume Exit_cmdFRMUndo_Click
'
'End Sub



Private Sub Form_Load()
    If gbPalmCommaDelimited Then
        Label57.Caption = "Receive from Palm (Comma)"
    Else
        Label57.Caption = "Receive from Palm (Tab)"
    End If
    'Me.InsideHeight = boxScreen.Height + FormFooter.Height + 250
    'Me.InsideWidth = boxScreen.Width + 500
    cboID.ColumnHidden = True
    SetupDate.ColumnHidden = True
    ChangeDate.ColumnHidden = True
    If gbClubExportOGFormatInExcel Then
        Label95.Caption = "Export ALL Order Guide to Excel .  Exported to K:\KMDS\Clubs\  "

    Else
        Label95.Caption = "Export ALL Order Guide to OO Excel.Exported to K:\KMDS\Clubs\  "
    End If
    If GetKMDSSettings("Fintech.Active", "False") = "True" Then
        FintechAccount.Visible = True
    Else
        FintechAccount.Visible = False
    End If




End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub cmdOrderGuides_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdOrderGuides_Click
    MsgBox ("Please maintain order guides thru the 'Create Club Order' Screen")
    Exit Sub
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ClubOrderGuideTemp"

    stLinkCriteria = "[ID]=" & Me![ID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    'cmdViewOrderGuide.Visible = True
    'cmdViewOrderGuide.SetFocus
    'cmdOrderGuides.Visible = False
    'Me.ClubOrderGuides.Visible = False

Exit_cmdOrderGuides_Click:
    Exit Sub

Err_cmdOrderGuides_Click:
    MsgBox Err.Description
    Resume Exit_cmdOrderGuides_Click

End Sub
Private Sub cmdViewOrderGuide_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdViewOrderGuide_Click
    Me.ClubOrderGuides.Requery
    Me.ClubOrderGuides.Visible = True

    'cmdOrderGuides.Visible = True
    txtName.SetFocus
    'cmdViewOrderGuide.Visible = False

Exit_cmdViewOrderGuide_Click:
    Exit Sub

Err_cmdViewOrderGuide_Click:
    MsgBox Err.Description
    Resume Exit_cmdViewOrderGuide_Click

End Sub
Private Sub cmdPrevInvoices_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrevInvoices_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Club"

    stLinkCriteria = "[ID]=" & Me![ID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdPrevInvoices_Click:
    Exit Sub

Err_cmdPrevInvoices_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrevInvoices_Click

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
Private Sub cmdAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddNew_Click


    DoCmd.GoToRecord , , acNewRec

Exit_cmdAddNew_Click:
    Exit Sub

Err_cmdAddNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNew_Click

End Sub
Private Sub Command69_Click()
If gcfHandleErrors Then On Error GoTo Err_Command69_Click

    Dim stDialStr As String
    Dim PrevCtl As Control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91
    Const ERR_CANTMOVE = 2483

    Set PrevCtl = Screen.PreviousControl

    If TypeOf PrevCtl Is TextBox Then
      stDialStr = IIf(varType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ListBox Then
      stDialStr = IIf(varType(PrevCtl) > V_NULL, PrevCtl, "")
    ElseIf TypeOf PrevCtl Is ComboBox Then
      stDialStr = IIf(varType(PrevCtl) > V_NULL, PrevCtl, "")
    Else
      stDialStr = ""
    End If

    Application.Run "utility.wlib_AutoDial", stDialStr

Exit_Command69_Click:
    Exit Sub

Err_Command69_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Or (Err = ERR_CANTMOVE) Then
      Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_Command69_Click

End Sub
Private Sub cmdExportToExcel_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdExportToExcel_Click

    Dim stDocName As String

    stDocName = "qryClubOrderGuidesForExport"
    DoCmd.OutputTo acReport, stDocName

Exit_cmdExportToExcel_Click:
    Exit Sub

Err_cmdExportToExcel_Click:
    MsgBox Err.Description
    Resume Exit_cmdExportToExcel_Click

End Sub

Private Sub cmdExportAllToExcel_Click()

    Dim stDocName As String
    Dim fso As FileSystemObject
    Dim strOutput As String
    Dim strEmail As String


    Dim f As Folder
    Dim ts As Object
    Dim fldr As String
    Dim coleMail As New Collection
    Dim strQuote As String
    strQuote = Chr(34)
    If gcfHandleErrors Then On Error GoTo ErrLog

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"
    Set fso = New FileSystemObject

    fldr = "\\" & gstrCashsysInsLoc & "\KMDS\Clubs\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If



    stDocName = "qryClubOrderGuidesForExport"
    '            qryClubOrderGuidesForExport
    DoCmd.GoToRecord , , acFirst

    Do While Not Me.Recordset.EOF
        If txtName > "" Then
            strOutput = fldr & txtName & ".xls"
            strOutput = "k:\clubs\" & txtName & ".xls"
            strEmail = "k:\clubs\" & txtName & ".xls"
            DoCmd.OutputTo acOutputQuery, stDocName, acFormatXLS, strOutput
            'DoCmd.OutputTo acOutputQuery, stDocName, acSpreadsheetTypeExcel9, strOutput
            If gbClubExportOGFormatInExcel Then
                WaitForExecCmd ("k:\ClubOGXL.exe " & strQuote & fldr & txtName & ".xls" & strQuote)
            Else
                WaitForExecCmd ("k:\clubOogl.exe " & strQuote & fldr & txtName & ".xls" & strQuote)
            End If
            If EMail > "" Then
                If fso.FileExists(strEmail) Then
                    If MsgBox(txtName & " eMail order guide?", vbYesNo) = vbYes Then
                        coleMail.Add strEmail
                        KMDSeMail EMail, coleMail, "Order guide from " & gstrStoreName, GetOOG()
                        coleMail.Remove (1)
                    End If
                End If
            End If
        End If
        DoCmd.GoToRecord , , acNext
    Loop
    Me.Undo
    DoCmd.GoToRecord , , acFirst

ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2105
        Me.Undo
        DoCmd.GoToRecord , , acFirst
        MsgBox "Finished!"
    Case Else
        MsgBox Err.Number & " " & Err.Description
    End Select
    Resume ExitSub

End Sub

Private Sub Command96_Click()
If gcfHandleErrors Then On Error GoTo Err_Command96_Click



    DoCmd.GoToRecord , , acFirst

Exit_Command96_Click:
    Exit Sub

Err_Command96_Click:
    MsgBox Err.Description
    Resume Exit_Command96_Click

End Sub
Private Sub Command97_Click()
If gcfHandleErrors Then On Error GoTo Err_Command97_Click


    DoCmd.GoToRecord , , acNext

Exit_Command97_Click:
    Exit Sub

Err_Command97_Click:
    MsgBox Err.Description
    Resume Exit_Command97_Click

End Sub


Private Sub cmdReprintInvoice_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdReprintInvoice_Click

    Dim stDocName As String

    stDocName = "ClubInvoiceLiquorAndNonLiquorForClubScreen"
    DoCmd.OpenReport stDocName, acViewPreview


Exit_cmdReprintInvoice_Click:
    Exit Sub

Err_cmdReprintInvoice_Click:
    MsgBox Err.Description
    Resume Exit_cmdReprintInvoice_Click

End Sub
Private Sub cmdMailOrderGuide_Click()

Dim stDocName As String
Dim fso As FileSystemObject

Dim f As Folder
Dim ts As Object
Dim fldr As String
Dim coleMail As New Collection
Dim strBody As String

    Dim strQuote As String
    strQuote = Chr$(34)

    If gcfHandleErrors Then On Error GoTo ErrLog

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"

    Set fso = New FileSystemObject

    fldr = "\\" & gstrCashsysInsLoc & "\KMDS\Clubs\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    If fso.FileExists(fldr & txtName & ".xls") Then
        fso.DeleteFile fldr & txtName & ".xls"
    End If

    stDocName = "qryClubOrderGuidesForExport"
    DoCmd.OutputTo acOutputQuery, stDocName, acFormatXLS, fldr & txtName & ".xls"
'    WaitForExecCmd ("c:\kmds\clubogxl.exe " & strQuote & fldr & Desc & ".xls" & strQuote)
    If gbClubExportOGFormatInExcel Then
        WaitForExecCmd ("k:\ClubOGXL.exe " & strQuote & fldr & txtName & ".xls" & strQuote)
    Else
        WaitForExecCmd ("k:\clubOogl.exe " & strQuote & fldr & txtName & ".xls" & strQuote)
    End If


    coleMail.Add fldr & txtName & ".xls"
    KMDSeMail EMail, coleMail, "Order guide from " & gstrStoreName, GetOOG()
    MsgBox "Complete"


ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "Mail Order guide"

    Resume ExitSub

End Sub


Private Sub cmdPreviousOrders_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPreviousOrders_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "ClubsPreviousOrderSub"
    DoCmd.OpenForm stDocName, acFormDS

Exit_cmdPreviousOrders_Click:
    Exit Sub

Err_cmdPreviousOrders_Click:
    MsgBox Err.Description
    Resume Exit_cmdPreviousOrders_Click

End Sub



Private Sub Licensee_Enter()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Dim nsaveID As Long
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog

    'save it now to get the defaults!

    If Me.NewRecord And Me.Dirty Then
        nsaveID = ID
        Me.Dirty = False
        Me.Refresh
        Set rs = Me.Recordset.Clone
        'nSaveID = ID

        rs.FindFirst "[ID] = " & Str(nsaveID)
        Me.Bookmark = rs.Bookmark
        Licensee.SetFocus
   End If
ExitLog:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitLog

End Sub

Private Sub optLicenseeSort_AfterUpdate()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case optLicenseeSort.Value
    Case 1
        'MsgBox "1 name"
        With cmbLicensee
            .ColumnWidths = "3.25 in;1 in;1.5  in;0.75 in;0.5 in"
            .RowSource = "SELECT  Licensee.TradeName, Licensee.PremisesNumber, Licensee.Street, Licensee.City, Licensee.Zip FROM Licensee ORDER BY Licensee.TradeName;"
            .BoundColumn = 2
        End With

    Case 2
        'MsgBox "2 number"
        With cmbLicensee
            .ColumnWidths = "1 in;3.25 in;1.5  in;0.75 in;0.5 in"
            .RowSource = "SELECT  Licensee.PremisesNumber,Licensee.TradeName, Licensee.Street, Licensee.City, Licensee.Zip FROM Licensee ORDER BY Licensee.PremisesNumber;"
            .BoundColumn = 1
        End With
    End Select
ExitLog:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitLog
End Sub
