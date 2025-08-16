Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True

End Sub

Private Sub Form_Load()
    gbRunMailChanges = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Cancel = FormLogin(glnCashierID, Me)
    ID.BackColor = glnColorLockedForAll
    Me!Name.BackColor = glnColorLockedForAll
    gbRunMailChanges = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

End Sub
