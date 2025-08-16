Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterInsert()
    Me.Requery

End Sub

Private Sub Form_AfterUpdate()
    'Me.Requery
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
ClubID = Forms![ClubOrderWebImport]![cmbLicensee]
End Sub

Private Sub Form_Current()
    If gcfHandleErrors Then On Error GoTo ErrLog

    LineID.ColumnWidth = -2
    ItemID.ColumnWidth = -2
    Combo16.ColumnWidth = -2
    Par.ColumnWidth = -2

'    QuotedPrice.ColumnWidth = -2
'    QuotedPrice.ColumnHidden = Not chkNonLiquor
'    'Par.ColumnHidden = Not chkMaintainParLevels

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Current"
    Resume ExitSub
End Sub
