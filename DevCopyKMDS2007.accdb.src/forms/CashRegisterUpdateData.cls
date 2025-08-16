Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    MailChanges False
    Cancel = FormLogin(glnCashierID, Me)
    Cancel = True


End Sub
Private Sub Form_Close()
    FormLogout glnCashierID, Me

End Sub
