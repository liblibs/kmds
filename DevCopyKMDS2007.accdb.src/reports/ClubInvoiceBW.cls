Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mnLinecount As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim ctl As Control
    mnLinecount = mnLinecount + 1
    For Each ctl In Detail.Controls
        If ctl.ControlType = acTextBox Then
            If mnLinecount > 1 Then
                ctl.BackColor = vbWhite
            Else
                ctl.BackColor = lblInvoiceDate.BackColor

            End If
        End If
    Next ctl

    If mnLinecount > 1 Then
        mnLinecount = 0
    End If

End Sub

Private Sub GroupFooter0_Format(Cancel As Integer, FormatCount As Integer)
   ' txtRepacks.Visible = Repack

End Sub

'Private Sub PageFooterSection_Format(Cancel As Integer, FormatCount As Integer)
'    If Page = Pages Then
'        PageFooterSection.Visible = True
'    Else
'        PageFooterSection.Visible = False
'    End If
'End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox "No data on report!"
    Cancel = True
End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    Me.OrderByOn = True
    Me.OrderBy = Forms("ClubOrderPrintDialog").InvoiceListSort

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "ClubPickList"
    Resume ExitSub
End Sub

Private Sub Report_Page()
    If Page = Pages Then
        Me.PageFooterSection.Visible = True
    Else
        Me.PageFooterSection.Visible = False
    End If
End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    txtStoreName = gstrStoreName
    txtCashier = CStr(glnCashierID) & " " & glsCashierName
End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)
    txtSignature.Visible = Not NonLiquor
    linSignature.Visible = Not NonLiquor
    If txtOrderNotes = "" Then
        txtOrderNotes.Visible = False
        'lblOrderNotes.Visible = False
    Else
        txtOrderNotes.Visible = True
        'lblOrderNotes.Visible = True
    End If
End Sub
