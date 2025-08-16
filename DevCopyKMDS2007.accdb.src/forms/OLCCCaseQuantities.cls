Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' Command0_Click
'
'------------------------------------------------------------
Private Sub Command0_Click()
Dim stDocName As String
On Error GoTo Command0_Click_Err


    SetOptions False
    stDocName = "C:\Npos\Case_Quantities.txt"

    DoCmd.OpenQuery "qryMinimumOrderQ_01_Delete", acViewNormal, acEdit

    DoCmd.TransferText acImportDelim, "Case_quantities Import Specification", "Case_quantities", stDocName, False
    DoCmd.RunSQL "delete * from Case_quantities where ((Case_quantities.CreateDate) Is Null)"



Command0_Click_Exit:
    Exit Sub

Command0_Click_Err:
    MsgBox Error$
    Resume Command0_Click_Exit

End Sub


'------------------------------------------------------------
' Command1_Click
'
'------------------------------------------------------------
Private Sub Command1_Click()
On Error GoTo Command1_Click_Err

    DoCmd.OpenQuery "qryMinimumOrderQ_01_Delete", acViewNormal, acEdit


Command1_Click_Exit:
    Exit Sub

Command1_Click_Err:
    MsgBox Error$
    Resume Command1_Click_Exit

End Sub
