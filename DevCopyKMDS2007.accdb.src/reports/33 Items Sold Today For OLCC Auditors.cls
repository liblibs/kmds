Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nCount5 As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    nCount5 = nCount5 + 1
    If nCount5 = 5 Then
        nCount5 = 0
        Line5.Visible = True
    Else
        Line5.Visible = False
    End If
End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)
    'Text69.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub

Private Sub GroupFooter5_Format(Cancel As Integer, FormatCount As Integer)
    'Text61.Visible = Forms![PrintInventoryReport]![chkCost]
    'Text66.Visible = Forms![PrintInventoryReport]![chkCost]
   ' Text67.Visible = Forms![PrintInventoryReport]![chkCost]

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    nCount5 = 0

End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox ("No data! Try Choosing a different range of parameters")
    Cancel = True

End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    Me.OrderBy = "ID"
    If Me.OpenArgs <> "" Then
        Me.OrderBy = Me.OpenArgs
    End If
    Me.OrderByOn = True
End Sub
