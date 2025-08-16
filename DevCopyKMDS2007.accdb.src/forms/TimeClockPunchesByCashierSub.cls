Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ChangeDate = Now
End Sub

Private Sub Form_Load()
    Wages.ColumnHidden = Not gbAllowHourlyWageView
    If Not gbAllowHourlyWageView Then
        Wages.ControlSource = ""
    End If

End Sub
