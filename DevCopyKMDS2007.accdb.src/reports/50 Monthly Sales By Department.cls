Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nCount5 As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'    Text66.Visible = Forms![PrintInventoryReport]![chkCost]
'    Text67.Visible = Forms![PrintInventoryReport]![chkCost]
'    Text51.Visible = Forms![PrintInventoryReport]![chkCost]
    nCount5 = nCount5 + 1
    If nCount5 = 5 Then
        nCount5 = 0
        Line5.Visible = True
    Else
        Line5.Visible = False
    End If
End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    'D1_Label.Visible = Forms![PrintInventoryReport]![chkCost]
    'D2_Label.Visible = Forms![PrintInventoryReport]![chkCost]
    nCount5 = 0

End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    Me.OrderBy = "ID"
    If Me.OpenArgs <> "" Then
        Me.OrderBy = Me.OpenArgs
    End If
    Me.OrderByOn = True
End Sub
