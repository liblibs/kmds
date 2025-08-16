Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim mbNoData As Boolean


Private Sub Report_Close()
    If mbNoData = True Then Exit Sub

    If MsgBox("Do you want to delete all of these UPC Codes?", vbYesNo) = vbYes Then
        If MsgBox("Are you sure?", vbYesNo) = vbYes Then
            QueryExecute "qryUPCItemsObsoleteDelete"
        End If
    End If


    End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox ("No data! There are no obsolete UPC Codes")
    Cancel = True
    mbNoData = True
End Sub

Private Sub Report_Open(Cancel As Integer)
    mbNoData = False

End Sub
