Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim nCount5 As Integer


Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    Cost.Visible = Forms![PrintInventoryReport]![chkCost]
    'ExtCost.Visible = Forms![PrintInventoryReport]![chkCost]
    Text66.Visible = Forms![PrintInventoryReport]![chkCost]
    Text67.Visible = Forms![PrintInventoryReport]![chkCost]
    'Text51.Visible = Forms![PrintInventoryReport]![chkCost]
    nCount5 = nCount5 + 1
    If nCount5 = 5 Then
        nCount5 = 0
        Line5.Visible = True
    Else
        Line5.Visible = False
    End If

End Sub

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)
    txtDepartmentCost.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)
    Text69.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub

Private Sub GroupFooter4_Format(Cancel As Integer, FormatCount As Integer)
    Text37.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub

Private Sub GroupFooter5_Format(Cancel As Integer, FormatCount As Integer)
    Text61.Visible = Forms![PrintInventoryReport]![chkCost]
    Text66.Visible = Forms![PrintInventoryReport]![chkCost]
    Text67.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    D1_Label.Visible = Forms![PrintInventoryReport]![chkCost]
    'D2_Label.Visible = Forms![PrintInventoryReport]![chkCost]
    nCount5 = 0
End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    Text42.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub
