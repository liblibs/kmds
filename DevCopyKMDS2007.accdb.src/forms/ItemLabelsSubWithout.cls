Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
'Me.FilterOn = False
'Me.FilterOn = True

End Sub

Private Sub Form_Load()
    Me.Filter = "((ItemLabels.Labels=0))"
    Me.FilterOn = True
    ID.ColumnWidth = 1500
    Labels.ColumnWidth = 500
End Sub

Private Sub Text25_DblClick(Cancel As Integer)

End Sub

Private Sub Labels_DblClick(Cancel As Integer)
    If Labels = 0 Then
        Labels = 1
    Else
        Labels = 0
    End If
End Sub
