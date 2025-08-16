Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Dim gbRunMailChanges As Boolean

Private Sub cmdfrmClose_Click()
    DoCmd.Close acForm, "ItemDataSheetSelect"

    DoCmd.Close
End Sub

Private Sub cmdOutput_Click()
'    ExportReport Me.Tag

    'DoCmd.OutputTo acOutputQuery, ItemDataSheetSub.Form.RecordSource, acFormatXLS, "c:\test.xls", True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    gbRunMailChanges = True
End Sub

Private Sub Form_Load()
    gbRunMailChanges = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Cancel = FormLogin(glnCashierID, Me)
    SetFormSize Me

End Sub
'
Private Sub Form_Resize()
If gcfHandleErrors Then On Error GoTo ErrLog
    ItemDataSheetSub.Width = Me.Form.InsideWidth - 250
    ItemDataSheetSub.Form.InsideHeight = Me.Form.InsideHeight - (Me.FormFooter.Height + 500) '- (Me.FormHeader.Height + Me.FormFooter.Height + 500)
ExitSub:
    Exit Sub
'
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Resize"
    Resume ExitSub
End Sub
'

Private Sub cmdPrintPreview_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdPrintPreview_Click

    Dim stDocName As String
    Me.Form.printform

'    stDocName = "ClubInvoice"
'    DoCmd.OpenReport stDocName, acPreview

Exit_cmdPrintPreview_Click:
    Exit Sub

Err_cmdPrintPreview_Click:
    MsgBox Err.Description
    Resume Exit_cmdPrintPreview_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then
        MailChanges (True)
    End If
End Sub
