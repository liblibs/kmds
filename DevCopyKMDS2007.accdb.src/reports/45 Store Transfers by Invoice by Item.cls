Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim nCount5 As Integer
Dim nc As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
   nCount5 = nCount5 + 1
    If nCount5 = 5 Then
        nCount5 = 0
        'Line5.Visible = True
    Else
        Line5.Visible = False
    End If
End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
    nc = nc + 1

End Sub

Private Sub GroupFooter0_Format(Cancel As Integer, FormatCount As Integer)
        nCount5 = 0

End Sub

Private Sub Report_NoData(Cancel As Integer)
    lblNoData.Visible = True
    'Cancel = True
End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
    Me.OrderBy = ""
    If Me.OpenArgs <> "" Then
        Me.OrderBy = Me.OpenArgs
    End If
    Me.OrderByOn = True
    nc = 0
    nCount5 = 0
End Sub
