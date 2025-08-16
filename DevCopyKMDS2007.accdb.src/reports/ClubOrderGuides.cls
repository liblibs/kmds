Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mnLinecount As Integer

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
Dim ctl As Control
    mnLinecount = mnLinecount + 1
    For Each ctl In Detail.Controls
        If ctl.ControlType = acTextBox Then
            If mnLinecount > 1 Then
                ctl.BackColor = vbWhite
                'ctl.FontBold = True
            Else
                ctl.BackColor = lblPageNumber.BackColor
                'ctl.FontBold = False

            End If
        End If
    Next ctl

    If mnLinecount > 1 Then
        mnLinecount = 0
    End If
    ItemID.BackColor = vbWhite
    Par.BackColor = vbWhite

End Sub


Private Sub Report_NoData(Cancel As Integer)
    Cancel = True
End Sub

Private Sub Report_Open(Cancel As Integer)

    If gcfHandleErrors Then On Error GoTo ErrLog
    Me.OrderByOn = True
    Me.OrderBy = Forms("ClubOrderPrintDialog").OrderGuideSort

ExitSub:
    Exit Sub
ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "ClubPickList"
    Resume ExitSub
End Sub
