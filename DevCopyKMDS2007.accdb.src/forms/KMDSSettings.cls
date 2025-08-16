Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    'try to prevent changing existing key names..
    If Me.NewRecord Then
        KeyName.Locked = False
    Else
        KeyName.Locked = True
    End If
End Sub
