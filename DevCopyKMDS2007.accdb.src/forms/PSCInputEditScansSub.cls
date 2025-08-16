Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
    If Not Me.NewRecord Then
        Item_Code.Locked = Locked
        Item_Description.Locked = Locked
        OnShelf.Locked = Locked
        Price.Locked = Locked
        Cost.Locked = Locked
        Locked.Locked = Locked
    End If
End Sub
