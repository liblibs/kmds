Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim nLcount As Integer

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
    If TransType <> "S" And TransType <> "C" Then
        InputCode.Visible = True
    Else
        InputCode.Visible = False
    End If
    nLcount = nLcount + 1
    If nLcount = 5 Then
        Line5.Visible = True
        nLcount = 0
    Else
        Line5.Visible = False
    End If
End Sub
