Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nCount5 As Integer

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

    nCount5 = 0

End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox ("No data! Try Choosing a different range of parameters")
    Cancel = True

End Sub

Private Sub Report_Open(Cancel As Integer)
    If gcfHandleErrors Then On Error Resume Next
'    Me.OrderBy = "SetupDate"
    If Me.OpenArgs <> "" Then
        Me.OrderBy = Me.OpenArgs
    End If
    Me.OrderByOn = True
End Sub
