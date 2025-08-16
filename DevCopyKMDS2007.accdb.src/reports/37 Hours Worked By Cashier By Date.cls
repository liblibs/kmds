Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim nCount5 As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'    nCount5 = nCount5 + 1
'    If nCount5 = 5 Then
'        Line5.Visible = True
'        nCount5 = 0
'    Else
'        Line5.Visible = False
'    End If

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    nCount5 = 0
End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox ("No data! Try Choosing a different range of parameters")
    Cancel = True

End Sub

Private Sub Report_Open(Cancel As Integer)
    If Not gbAllowHourlyWageView Then
        lblWages.Visible = False
        Wages.Visible = False
        DTotalWages.Visible = False
        GTotalWages.Visible = False
    End If
End Sub
