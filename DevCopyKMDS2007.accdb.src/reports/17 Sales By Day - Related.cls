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
