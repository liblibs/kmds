Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command14_Click()
If gcfHandleErrors Then On Error GoTo Err_Command14_Click

    Dim stDocName As String

    stDocName = "qryAverageDaySupplyAppend"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_Command14_Click:
    Exit Sub

Err_Command14_Click:
    MsgBox Err.Description
    Resume Exit_Command14_Click

End Sub

Private Sub GroupLevel_DblClick(Cancel As Integer)
Command14_Click
Me.Requery
Me.Refresh
End Sub
