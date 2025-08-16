Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    PunchDate = Me.Parent.PunchDate
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If HoursWorked < 0 Then
        If Not gcfHandleErrors Then Debug.Print Punch2
        MsgBox "Hours worked can not be negative!"
        If Punch2 < #11:59:59 AM# Then
            Punch2 = DateAdd("h", 12, Punch2)
        End If
        Cancel = True
        Punch2.SetFocus
    End If
    ChangeDate = Now
End Sub

Private Sub Form_Load()
    Wages.ColumnHidden = Not gbAllowHourlyWageView
    If Not gbAllowHourlyWageView Then
        Wages.ControlSource = ""
    End If
End Sub
