Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    If IsNull(FAX) Or FAX = "" Then
        FAX.Visible = False
    Else
        FAX.Visible = True
    End If
End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    Me.OrderByOn = True
    Me.OrderBy = Forms("PurchaseOrderChoices").Phone3

End Sub

Private Sub Report_Page()
    If gcfHandleErrors Then On Error Resume Next
    If Page = txtPages Then
        Me.PageFooterSection.Visible = False
    Else
        Me.PageFooterSection.Visible = True
    End If
End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)
    If Forms![PurchaseOrderChoices]![optChoices] > 1 Then
        txtPosted = "Posted " & Format(Now, "mm/dd/yy hh:mm")
        txtPosted.Visible = True

    Else
        txtPosted.Visible = False
    End If

End Sub
