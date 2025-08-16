Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim mbDropsOrLoansFound As Boolean

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

    GroupFooter2.Visible = True
    If gcfHandleErrors Then On Error Resume Next ' when there is no data an error is generated
    Select Case txtGroupLevel
    Case 1
        txtGroupLevelDesc = "Sales"
    Case 2
        mbDropsOrLoansFound = True
        txtGroupLevelDesc = "Drawer Counts"
    Case Else
        GroupFooter2.Visible = False
    End Select

End Sub

Private Sub Report_NoData(Cancel As Integer)

    lblNoData.Visible = True
    Cancel = True
End Sub

Private Sub Report_Open(Cancel As Integer)
    mbDropsOrLoansFound = False
    GetAllRegistersDetailFiles
End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    Me("D1 Grand Total Sum").Visible = mbDropsOrLoansFound
    Me("D2 Grand Total Sum").Visible = mbDropsOrLoansFound
    Me("D3 Grand Total Sum").Visible = mbDropsOrLoansFound
    Me("D4 Grand Total Sum").Visible = mbDropsOrLoansFound
    Me("TotalP Grand Total Sum").Visible = mbDropsOrLoansFound

    If mbDropsOrLoansFound Then
        If Not gcfHandleErrors Then Debug.Print Me("TotalP Grand Total Sum")
        If Me("TotalP Grand Total Sum") > 0 Then
            Label44.Caption = "Total OVER"
        End If
        If Me("TotalP Grand Total Sum") < 0 Then
            Label44.Caption = "Total SHORT"
        End If
        If Me("TotalP Grand Total Sum") = 0 Then
            Label44.Caption = "BALANCED!"
        End If
    End If
End Sub
