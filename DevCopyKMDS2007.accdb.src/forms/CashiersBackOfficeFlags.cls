Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    'Cancel = FormLogin(glnCashierID, Me)
    'SetFormSize Me

    ID.BackColor = glnColorLockedForAll
    Me!Name.BackColor = glnColorLockedForAll

End Sub
