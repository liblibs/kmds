Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
    Me.Parent("LabelMenuOption").Caption = MenuOption
    Me.Parent("LabelMenuOption").Tag = Argument
    Me.Parent("CashierRightsCashierListSub").Requery
    Me.Parent("cmdSetAllowedForThisProgram").Caption = "Se&t ALL Allowed Flags for " & MenuOption ' .Column(5)
    Me.Parent("cmdSetAllowedForThisProgram").Tag = Argument ' .Column(5)
End Sub
