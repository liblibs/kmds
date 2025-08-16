Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command0_Click()
If gcfHandleErrors Then On Error GoTo Err_Command0_Click

    Dim stDocName As String
    SetOptions False
    If Me.Dirty Then Me.Dirty = False
    stDocName = "qryPalm"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

    DoCmd.TransferText acExportDelim, "Palm Export Specification", "aPalm", "\Program Files\PTS\TracerPlus 2.8\Data\CBS\ToPDA\s03_data.txt", False
    SetOptions True


Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub
