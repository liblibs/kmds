Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'------------------------------------------------------------
' cmdSBLCImportSpecification_Click
'
'------------------------------------------------------------

Private Sub Form_Load()



End Sub

Private Sub Form_Open(Cancel As Integer)
Dim txtstr As String
Dim txtDirSearch As String

If gcfHandleErrors Then On Error GoTo ErrLog
txtDirSearch = "C:\kmds\sblc\sblc" & Format(Now, "yyyy-mm-dd") & "*.txt"

txtstr = Dir(txtDirSearch)
If txtstr <> "" Then
    lstSBLC.AddItem txtstr
    Do While txtstr <> ""
        txtstr = Dir
        lstSBLC.AddItem txtstr
    Loop
Else
    MsgBox "There are no batches found matching " & vbCr & txtDirSearch
    Cancel = True
End If
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description & " Form_Load"
    Resume ExitSub
End Sub

Private Sub lstSBLC_DblClick(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    DoCmd.TransferText acImportDelim, "SBLCImportSpecification", "ScanBatchLineCounts", "C:\KMDS\SBLC\" & lstSBLC, True
    MsgBox "This batch has been restored!"
    DoCmd.Close
    DoCmd.OpenForm "ScanBatch"
ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description & " lstSBLC_DblClick"
    Resume ExitSub

End Sub
