Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean

Private Sub Code_DblClick(Cancel As Integer)
    OpenItemForm Code
End Sub

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True
End Sub

Private Sub Form_Load()
    Me.OrderBy = "Line"
    Me.OrderByOn = True
    gbRunMailChanges = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False

End Sub
