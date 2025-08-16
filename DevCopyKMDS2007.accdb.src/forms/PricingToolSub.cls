Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub Form_DblClick(Cancel As Integer)
'    Me.NewPrice = Me.Price
'    Form.Dirty = False
'End Sub

Private Sub NewPrice_AfterUpdate()
    NewPrice = Round(NewPrice, 2)
    If Me.NewPrice > 0 Then

        Me.NewGP = ((Me.NewPrice - Me.Cost) / Me.NewPrice)
    Else
        Me.NewGP = 0
    End If
End Sub

Private Sub NewPrice_Click()
    Me.NewPrice = Me.Price
    If Me.NewPrice > 0 Then
        Me.NewGP = ((Me.NewPrice - Me.Cost) / Me.NewPrice)
    Else
        Me.NewGP = 0
    End If
    Form.Dirty = False

End Sub
