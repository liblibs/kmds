Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    Me.Caption = "Items on order for " & Me.OpenArgs
    Me.DatasheetBackColor = vbYellow
End Sub
