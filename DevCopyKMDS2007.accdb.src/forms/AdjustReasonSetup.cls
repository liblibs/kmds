Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

Private Sub Form_Current()
    gstrAdjustReason = Nz(Reason, "Not given")
End Sub
Private Sub List4_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![List4], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub

Private Sub List4_Click()
cmdfrmClose_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
    'MsgBox Me.DefaultView


End Sub
