Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Timer()
    With Label1
        .ForeColor = (IIf(.ForeColor = 0, 255, 0))
    End With
    With Text0
        .BackColor = (IIf(.BackColor = 0, 255, 0))
    End With

End Sub
