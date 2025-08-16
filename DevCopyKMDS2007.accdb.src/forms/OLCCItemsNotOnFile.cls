Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private Sub cmdDownload_Click()
    Dim strSQL As String
    strSQL = "You only need to do this when the downloaded column is prior to the current month" & vbCrLf
    strSQL = strSQL & "Do you want to continue?"
    If MsgBox(strSQL, vbYesNo) = vbYes Then
        GetOLCCItemFiles
        MsgBox "Please reload this program"
        DoCmd.Close acForm, Me.Name
    End If
End Sub


Private Sub Form_Current()
    On Error Resume Next
    If Month(Now) <> Month(txtSetupDate) Then
        cmdDownload.Visible = True
        cmdDownload.SetFocus
        txtSetupDate.Visible = False
    Else
        txtSetupDate.Visible = True
        txtSetupDate.SetFocus
        cmdDownload.Visible = False
    End If


End Sub

Private Sub Form_DblClick(Cancel As Integer)
        cmdDownload.Visible = True

End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.InsideHeight = boxScreen.Height + 100 'FormFooter.Height + 100
    Me.InsideWidth = boxScreen.Width + 100

End Sub

Private Sub txtSetupDate_DblClick(Cancel As Integer)
        cmdDownload.Visible = True

End Sub
