Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdSave_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdSave_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSave_Click:
    Exit Sub

Err_cmdSave_Click:
    MsgBox Err.Description
    Resume Exit_cmdSave_Click

End Sub
Private Sub Command9_Click()
If gcfHandleErrors Then On Error GoTo Err_Command9_Click


    DoCmd.Close

Exit_Command9_Click:
    Exit Sub

Err_Command9_Click:
    MsgBox Err.Description
    Resume Exit_Command9_Click

End Sub
