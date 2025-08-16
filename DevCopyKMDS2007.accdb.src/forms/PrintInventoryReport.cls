Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboCategoryIDBegin_AfterUpdate()
    cboCategoryIDEnd = cboCategoryIDBegin
End Sub

Private Sub cboClubIDBegin_AfterUpdate()
    cboClubIDEnd = cboClubIDBegin
End Sub

Private Sub cboItemIDBegin_AfterUpdate()
    cboItemIDEnd = cboItemIDBegin
End Sub

Private Sub cboLocationBegin_AfterUpdate()
    cboLocationEnd = cboLocationBegin
End Sub

Private Sub cboPriceGroupIDBegin_AfterUpdate()
    cboPriceGroupIDEnd = cboPriceGroupIDBegin
End Sub


Private Sub cboSalesDepartmentIDBegin_AfterUpdate()
    cboSalesDepartmentIDEnd = cboSalesDepartmentIDBegin
End Sub

Private Sub cboVendorIDBegin_AfterUpdate()
    cboVendorIDEnd = cboVendorIDBegin
End Sub

Private Sub cmbOrderBy_DblClick(Cancel As Integer)
    GenReport acPreview
End Sub

Private Sub cmdArchiveReport_Click()
    GenArchiveReport
End Sub

Private Sub cmdChart_Click()
    GenChart False
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

Private Sub cmdExportExcel_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdOutputToFile_Click

    GenReportOutputToExcel
Exit_cmdOutputToFile_Click:
    Exit Sub

Err_cmdOutputToFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdOutputToFile_Click

End Sub

Private Sub cmdPrint_Click()
    If txtReportName = "67 Items Restock List" Then ArchiveReport "qryEOD126ItemRestockList", glsCashierName

    GenReport acNormal
End Sub

Private Sub cmdPrintPreview_Click()
    GenReport acPreview
End Sub


Private Sub GenArchiveReport()

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    Dim stOrderBy As String
    UpdateDateTextBoxes
    stDocName = txtReportName
    If cmbOrderBy.Visible Then
        stOrderBy = Nz(cmbOrderBy, "")
    End If
    ArchiveReport stDocName, glsCashierName

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub

End Sub

Private Sub GenReport(DisplayMode As Integer)

If gcfHandleErrors Then On Error GoTo ErrLog
If Not gcfHandleErrors Then On Error GoTo ErrLog
    'MsgBox "GenReport"
    Dim stDocName As String
    Dim stOrderBy As String
    UpdateDateTextBoxes

    'preprocessing is required by some reports to get data
    Select Case txtReportName
    Case "24 Club History By Bottle By Month"   'run the append queries..
        SetOptions False
        If gcfHandleErrors Then On Error Resume Next
        DoCmd.DeleteObject acTable, "ClubHistoryByBottleByMonth"
        If gcfHandleErrors Then On Error GoTo ErrLog
        DoCmd.OpenQuery "qryClubHistoryByBottleByMonthMakeTable"
        Add12MonthsToClubHistoryTable
        SetOptions True
    Case "71 Bottle History By Club Month"   'run the append queries..
        SetOptions False
        If gcfHandleErrors Then On Error Resume Next
        DoCmd.DeleteObject acTable, "ClubHistoryByBottleByMonth"
        If gcfHandleErrors Then On Error GoTo ErrLog
        DoCmd.OpenQuery "qryClubHistoryByBottleByMonthMakeTable"
        Add12MonthsToClubHistoryTable
        SetOptions True
    Case Else

    End Select

    stDocName = txtReportName
    If cmbOrderBy.Visible Then
        stOrderBy = Nz(cmbOrderBy, "")
    End If

    DoCmd.OpenReport stDocName, DisplayMode, , , , stOrderBy
ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501
        'report cancel error
    Case Else
        MsgBox Err.Number & " " & Err.Description
    End Select
    Resume ExitSub

End Sub

Private Sub GenChart(PrintPrev As Boolean)

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    Dim stOrderBy As String
    UpdateDateTextBoxes
    stDocName = lstReportName.Column(10)

    DoCmd.OpenQuery stDocName, acViewPivotChart, acReadOnly
    If PrintPrev Then DoCmd.OpenQuery stDocName, acViewPreview

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub

End Sub

Private Sub GenReportOutputTo()

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    Dim stOrderBy As String
    UpdateDateTextBoxes
    stDocName = txtReportName

    If cmbOrderBy.Visible Then
        stOrderBy = Nz(cmbOrderBy, "")
    End If
    DoCmd.OutputTo acReport, stDocName
ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub

End Sub


Private Sub GenReportOutputToExcel()

If gcfHandleErrors Then On Error GoTo ErrLog

    Dim stDocName As String
    Dim stOrderBy As String
    UpdateDateTextBoxes
    stDocName = txtQueryName
    DoCmd.OutputTo acOutputQuery, stDocName, acFormatXLS, , True

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub

End Sub


Private Sub UpdateDateTextBoxes()
    txtDateRangeFrom = Format(dtpDateRangeFrom, "mm/dd/yyyy 00:00:00")
    txtDateRangeTo = Format(dtpDateRangeTo, "mm/dd/yyyy 23:59:59")
End Sub

Private Sub Detail_DblClick(Cancel As Integer)
    Description.Visible = True
End Sub

Private Sub dtpDateRangeFrom_Updated(Code As Integer)
    'dtpDateRangeTo.Value = dtpDateRangeFrom.Value
    UpdateDateTextBoxes
End Sub

Private Sub dtpDateRangeTo_Updated(Code As Integer)
    UpdateDateTextBoxes
End Sub

Private Sub Form_Current()
    lstReportName = txtReportName

    cboSalesDepartmentIDBegin.Visible = lstReportName.Column(2)
    cboSalesDepartmentIDEnd.Visible = lstReportName.Column(2)

    cboCategoryIDBegin.Visible = lstReportName.Column(17)
    cboCategoryIDEnd.Visible = lstReportName.Column(17)


    cboVendorIDBegin.Visible = lstReportName.Column(3)
    cboVendorIDEnd.Visible = lstReportName.Column(3)

    cboLocationBegin.Visible = lstReportName.Column(4)
    cboLocationEnd.Visible = lstReportName.Column(4)

    cboPriceGroupIDBegin.Visible = lstReportName.Column(11)
    cboPriceGroupIDEnd.Visible = lstReportName.Column(11)

    cboItemIDBegin.Visible = lstReportName.Column(5)
    cboItemIDEnd.Visible = lstReportName.Column(5)


    dtpDateRangeFrom.Visible = lstReportName.Column(6)
    dtpDateRangeTo.Visible = lstReportName.Column(6)
    framDates.Visible = lstReportName.Column(6)

    cboClubIDBegin.Visible = lstReportName.Column(16)
    cboClubIDEnd.Visible = lstReportName.Column(16)

    txtNumericRangeFrom.Visible = lstReportName.Column(12)
    txtNumericRangeTo.Visible = lstReportName.Column(12)
    Label60.Caption = Nz(lstReportName.Column(13), "")


    chkPrintOnhands.Visible = lstReportName.Column(7)
    chkCost.Visible = lstReportName.Column(8)
    chkPrintZeroOnhands.Visible = lstReportName.Column(9)

    cmbOrderBy = ""

    cmbOrderBy.Requery
    cmbOrderBy.Visible = cmbOrderBy.ListCount

    If lstReportName.Column(10) = "" Then
        cmdChart.Visible = False
        cmdPrintChart.Visible = False
    Else
        cmdChart.Visible = True
        cmdPrintChart.Visible = True
    End If

    Description.Visible = Not IsNull(Description)
    cmdExportExcel.Visible = Not IsNull(txtQueryName)
    If Me.OpenArgs <> "" Then 'called by the show detail button in item setup screen
        Select Case Left(txtReportName, 2)
        Case "18"
            cboItemIDBegin = Me.OpenArgs
            cboItemIDEnd = cboItemIDBegin
            dtpDateRangeFrom = CDate(Month(Date) & "/01/" & Year(Date))
            dtpDateRangeTo = Now
        Case "24"
            cboClubIDBegin = Me.OpenArgs
            cboClubIDEnd = cboClubIDBegin
            'dtpDateRangeFrom = CDate(Month(Date) & "/01/" & Year(Date))
            dtpDateRangeFrom = CDate(Format("01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
            dtpDateRangeTo = Now
        Case "26"
            cboClubIDBegin = Me.OpenArgs
            cboClubIDEnd = cboClubIDBegin
            'dtpDateRangeFrom = CDate(Month(Date) & "/01/" & Year(Date))
            dtpDateRangeFrom = CDate(Format("01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
            dtpDateRangeTo = Now
        Case "52"
            cboVendorIDBegin = Me.OpenArgs
            cboVendorIDEnd = cboVendorIDBegin

            With cboSalesDepartmentIDBegin
                cboSalesDepartmentIDBegin = .ItemData(0)
                cboSalesDepartmentIDEnd = (.ItemData(.ListCount - 1))
            End With

            With cboPriceGroupIDBegin
                cboPriceGroupIDBegin = .ItemData(0)
                cboPriceGroupIDEnd = (.ItemData(.ListCount - 1))
            End With
            dtpDateRangeFrom = CDate(Format("01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
            dtpDateRangeTo = Now
        End Select

    End If
End Sub


Private Sub Form_Load()

    If Me.OpenArgs <> "" Then 'called by the show detail button in item setup screen
        lstReportName.Visible = False
'        Select Case Left(txtReportName, 2)
'        Case "18"
'            cboItemIDBegin = Me.OpenArgs
'            cboItemIDEnd = cboItemIDBegin
'            dtpDateRangeFrom = CDate(Month(Date) & "/01/" & Year(Date))
'            dtpDateRangeTo = Now
'        Case "24"
'            cboClubIDBegin = Me.OpenArgs
'            cboClubIDEnd = cboClubIDBegin
'            'dtpDateRangeFrom = CDate(Month(Date) & "/01/" & Year(Date))
'            dtpDateRangeFrom = CDate(Format("01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
'           dtpDateRangeTo = Now
'        Case "52"
'            cboVendorIDBegin = Me.OpenArgs
'            cboVendorIDEnd = cboVendorIDBegin
'
'            With cboSalesDepartmentIDBegin
'                cboSalesDepartmentIDBegin = .ItemData(0)
'                cboSalesDepartmentIDEnd = (.ItemData(.ListCount - 1))
'            End With
'
'            With cboPriceGroupIDBegin
'                cboPriceGroupIDBegin = .ItemData(0)
'                cboPriceGroupIDEnd = (.ItemData(.ListCount - 1))
'            End With
'            dtpDateRangeFrom = CDate(Format("01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
'            dtpDateRangeTo = CDate(Format("12/31/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
'
'        End Select

    End If

    SetFormSize Me
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub framDates_BeforeUpdate(Cancel As Integer)
    Dim dtTemp As Date

    dtpDateRangeTo = Now
    Select Case framDates.Value
    Case 1 'today
        dtpDateRangeFrom = Now

    Case 2 'yesterday
        dtpDateRangeFrom = DateAdd("d", -1, Now)
    Case 3 'this week
        dtpDateRangeFrom = Now - (Weekday(Now) - 1) 'DateAdd("ww", -1, Now)
    Case 4 'this month
        dtpDateRangeFrom = CDate(Format(Now, "mm") & "/01/" & Format(Now, "yyyy"))
    Case 5 'this year
        dtpDateRangeFrom = CDate("01/01/" & Format(Now, "yyyy"))
    Case 6 'last week
        dtTemp = DateAdd("d", -7, Now)
        dtpDateRangeFrom = dtTemp - (Weekday(dtTemp) - 1)
        dtpDateRangeTo = DateAdd("d", 6, dtpDateRangeFrom)
    Case 7 'last month
        dtpDateRangeFrom = CDate(Format(Now, "mm") & "/01/" & Format(Now, "yyyy"))
        dtpDateRangeFrom = DateAdd("m", -1, dtpDateRangeFrom)
        dtpDateRangeTo = DateAdd("d", -1, CDate(Format(Now, "mm") & "/01/" & Format(Now, "yyyy")))
    Case 8 'last year
        dtpDateRangeFrom = CDate(Format("01/01/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))
        dtpDateRangeTo = CDate(Format("12/31/" & Format(DateAdd("yyyy", -1, Now), "yyyy")))

    Case Else


    End Select

End Sub

Private Sub lstReportName_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ReportName] = '" & Me![lstReportName] & "'"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    Me.Refresh
End Sub



Private Sub lstReportName_DblClick(Cancel As Integer)
    GenReport acPreview
End Sub
Private Sub cmdMail_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdMail_Click

    Dim stDocName As String
    UpdateDateTextBoxes
    stDocName = txtReportName

    DoCmd.SendObject acReport, stDocName

Exit_cmdMail_Click:
    Exit Sub

Err_cmdMail_Click:
    MsgBox Err.Description
    Resume Exit_cmdMail_Click

End Sub

Private Sub cmdPrintChart_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrintChart_Click

    GenChart True

Exit_cmdPrintChart_Click:
    Exit Sub

Err_cmdPrintChart_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintChart_Click

End Sub
Private Sub cmdOutputToFile_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdOutputToFile_Click

    GenReportOutputTo

Exit_cmdOutputToFile_Click:
    Exit Sub

Err_cmdOutputToFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdOutputToFile_Click

End Sub
