Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mnLinecount As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim ctl As Control
    If Not PoAtCost Then
        mnLinecount = mnLinecount + 1
        For Each ctl In Detail.Controls
            If ctl.ControlType = acTextBox Then
                If mnLinecount > 1 Then
                    ctl.BackColor = vbWhite
                Else
                    ctl.BackColor = lblQuantity.BackColor

                End If
            End If
        Next ctl

        If mnLinecount > 1 Then
            mnLinecount = 0
        End If
    End If
End Sub

Private Sub GroupFooter0_Format(Cancel As Integer, FormatCount As Integer)
    'txtRepacks.Visible = Repack

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    'If Forms![PurchaseOrderChoices]![optChoices] > 1 Then
        txtPosted = "Posted " & Format(txtPostDate, "mm/dd/yy hh:mm")
        txtPosted.Visible = True

    'Else
    '    txtPosted.Visible = False
    'End If
    If IsNull(FAX) Or FAX = "" Then
        FAX.Visible = False
    Else
        FAX.Visible = True
    End If

End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox "No data on report!"
    Cancel = True
End Sub

Private Sub Report_Open(Cancel As Integer)
'Dim qdftemp As DAO.QueryDef
'    Set qdftemp = CurrentDb().QueryDefs("qryPurchaseOrder")
'    qdftemp.Parameters("parPurchaseOrderID") = Forms!PurchaseOrdersOnFile.PurchaseOrderID
'    Me.RecordSource = qdftemp.SQL ' .OpenRecordset

    If gcfHandleErrors Then On Error Resume Next
    Me.OrderByOn = True
    Me.OrderBy = Forms("PurchaseOrderChoices").Phone3

If Me.OpenArgs = "NoPrice" Then
    Price.Visible = False
    ExtPrice.Visible = False
    lblPrice.Visible = False
    lblExtPrice.Visible = False
    linPrice.Visible = False
    linExtPrice.Visible = False
    txtBackGround.Width = txtBackGround.Width - (lblPrice.Width + lblExtPrice.Width)
    lblExtPriceHeader.Visible = False
    txtExtPriceHeader.Visible = False

    GPercent.Visible = False
    GPLabel.Caption = ""
    txtExtGPHeader.Visible = False
End If

End Sub

Private Sub Report_Page()
'    If Page = txtPages Then
'        Me.PageFooterSection.Visible = False
'    Else
'        Me.PageFooterSection.Visible = True
'    End If
End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    txtStoreName = gstrStoreName
    txtCashier = CStr(glnCashierID) & " " & glsCashierName
End Sub
