Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Combo6_AfterUpdate()
End Sub

Private Sub cmbPunchDate_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[PunchDate] = #" & Me![cmbPunchDate] & "#"
    If Not rs.EOF Then
        Me.Bookmark = rs.Bookmark
    Else
        MsgBox ("Records not found")
        cmbPunchDate = cmbPunchDate.OldValue
    End If

End Sub


Private Sub cmdHourly_Click()
    DoCmd.OpenForm "PrintInventoryReport", acNormal, , "([InventoryReportParameters].[ReportName]='36 Hours Worked By Date By Cashier')"
End Sub


Private Sub cmdHourly37_Click()
    DoCmd.OpenForm "PrintInventoryReport", acNormal, , "([InventoryReportParameters].[ReportName]='37 Hours Worked By Cashier By Date')"
End Sub

Private Sub Form_Current()
    If Me.NewRecord Then
    Else
'        If PunchDate = Date Then
'            MsgBox ("Changing today's records; End of day will replace these records")
'        End If
'
        cmbPunchDate = PunchDate
    End If
End Sub

Private Sub cmdAddNew_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdAddNew_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    MsgBox "Please note: any dates added in the future may be overwritten by a user punching in."


    stDocName = "TimeClockPunches"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormAdd
    DoCmd.Close acForm, Me.Name

Exit_cmdAddNew_Click:
    Exit Sub

Err_cmdAddNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdAddNew_Click

End Sub

Private Sub Form_Load()
    TotalWages.Visible = gbAllowHourlyWageView
    Me.InsideHeight = Box5.Height + 100
    Me.InsideWidth = Box5.Width + 500

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub
