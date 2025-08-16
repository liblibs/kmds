Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
    gbRunMailChanges = True
End Sub

Private Sub Form_Current()
    '[Forms]![ItemDataSheet]![ItemDataSheetSub]
    If gcfHandleErrors Then On Error Resume Next
    If Not gcfHandleErrors Then On Error Resume Next
    [Forms]![Cashier]![CashiersAndProgramsDatasheet].Form.Filter = "((Lookup_ProgramName.ItemText='" & cmbProgramName.Column(5) & "'))"
    [Forms]![Cashier]![CashiersAndProgramsDatasheet].Form.FilterOn = True
    [Forms]![Cashier]![lblCurrentProgramName].Caption = cmbProgramName.Column(5)
    [Forms]![Cashier]![cmdSetAllowedForThisProgram].Caption = "Se&t ALL Allowed Flags for " & cmbProgramName ' .Column(5)
    [Forms]![Cashier]![cmdSetAllowedForThisProgram].Tag = cmbProgramName  ' .Column(5)

    '((Lookup_ProgramName.ItemText="1. Cash Register"))
    '((Lookup_CashierID.Name="Steve Kelly"))
End Sub

Private Sub Form_Load()
    gbRunMailChanges = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    'If gcfHandleErrors Then On Error Resume Next
    Me.OrderBy = "Lookup_cmbProgramName.It"
    Me.OrderByOn = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbRunMailChanges Then MailChanges False
End Sub
