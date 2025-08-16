Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

    GroupFooter2.Visible = True
    If gcfHandleErrors Then On Error Resume Next ' when there is no data an error is generated
    Select Case txtGroupLevel
    Case 1
        txtGroupLevelDesc = "Sales"
    Case 2
        txtGroupLevelDesc = "Credits"
    Case Else
        GroupFooter2.Visible = False
    End Select

End Sub

Private Sub Report_NoData(Cancel As Integer)
      Cancel = True

    lblNoData.Visible = True
End Sub
