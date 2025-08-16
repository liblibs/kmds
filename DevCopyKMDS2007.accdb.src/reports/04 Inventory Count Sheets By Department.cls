Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim nCount5 As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    OnTheShelf.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    Onhand.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    OnReserve.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    nCount5 = nCount5 + 1
    If nCount5 = 5 Then
        nCount5 = 0
        Line5.Visible = True
    Else
        Line5.Visible = False
    End If

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    Text95.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    Text96.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    Text97.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    nCount5 = 0
End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    Me.OrderBy = "Description"
    If Me.OpenArgs <> "" Then
        Me.OrderBy = Me.OpenArgs
    End If
    Me.OrderByOn = True
End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    Text41.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    Text93.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]
    Text94.Visible = Forms![PrintInventoryReport]![chkPrintOnhands]

End Sub
