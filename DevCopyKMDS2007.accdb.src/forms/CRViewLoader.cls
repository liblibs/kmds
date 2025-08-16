Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command0_Click()
If gcfHandleErrors Then On Error GoTo Err_Command0_Click

    Dim stAppName As String

    stAppName = "\KMDS\CRView.exe"
    Call Shell(stAppName, 1)

Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub

Private Sub Form_Open(Cancel As Integer)
Command0_Click
Cancel = True
End Sub
