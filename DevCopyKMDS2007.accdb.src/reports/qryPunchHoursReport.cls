Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
    lblWages.Visible = gbAllowHourlyWageView
    TotalWages.Visible = gbAllowHourlyWageView
    CashierWages.Visible = gbAllowHourlyWageView
End Sub
