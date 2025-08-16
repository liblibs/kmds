Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command8_Click()
    Select Case Len(ID)
    Case 6
    Case 8

    Case 10
        txtUPCA = BuildUPCAGlyph("0" & CheckDigit11(ID))
   Case 12
        txtUPCA = BuildUPCAGlyph(ID)
    End Select
End Sub
