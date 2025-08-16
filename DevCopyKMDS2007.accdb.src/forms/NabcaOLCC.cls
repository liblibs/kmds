Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExportMonthendInventory_Click()
On Error GoTo cmdExportMonthendInventory_Click_Err
    SetOptions False

    NabcaOLCCInventFile CLng(glnAgentNumber), False

cmdExportMonthendInventory_Click_Exit:
    SetOptions True
    Exit Sub

cmdExportMonthendInventory_Click_Err:
    MsgBox Error$
    Resume cmdExportMonthendInventory_Click_Exit
End Sub

'------------------------------------------------------------
' cmdItemExample_Click
'
'------------------------------------------------------------
Private Sub cmdItemExample_Click()
On Error GoTo cmdItemExample_Click_Err
    SetOptions False

    GetNabcaOLCCItemFiles
    Me.Requery
    NabcaOlccItemSub.Requery

cmdItemExample_Click_Exit:
    SetOptions True
    Exit Sub

cmdItemExample_Click_Err:
    MsgBox Error$
    Resume cmdItemExample_Click_Exit

End Sub
