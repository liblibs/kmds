Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True

End Sub

Private Sub Form_Current()
   cmdfrmUndo.Visible = Me.NewRecord

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    cmdfrmUndo.Visible = True

End Sub

Private Sub Form_Load()
    gbRunMailChanges = False
    'Me.InsideHeight = boxScreen.Height + 100
    'Me.InsideWidth = boxScreen.Width + 500


End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

End Sub
