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
On Error GoTo Command0_Click_Err
FinTechExport



Command0_Click_Exit:
    Exit Sub

Command0_Click_Err:
    MsgBox Error$
    Resume Command0_Click_Exit
End Sub
