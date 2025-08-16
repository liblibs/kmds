Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrMonth

Private Sub MonExit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Apr_Enter()
    mstrMonth = "Apr"
End Sub

Private Sub Apr_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Aug_Enter()
    mstrMonth = "Aug"
End Sub

Private Sub Aug_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub cmdReset_Click()
Month1 = (1 / 13)
Month2 = (1 / 13)
Month3 = (1 / 13)
Month4 = (1 / 13)
Month5 = (1 / 13)
Month6 = (1 / 13)
Month7 = (1 / 13)
Month8 = (1 / 13)
Month9 = (1 / 13)
Month10 = (1 / 13)
Month11 = (1 / 13)
Month12 = (1 / 13)
CurrentMonth = (1 / 13)
CheckDirty


End Sub

Private Sub CurrentMonth_Exit(Cancel As Integer)
    CheckDirty
End Sub

Private Sub CurrMM_BeforeUpdate(Cancel As Integer)
'    If Me.Form.Dirty Then
'        'Me.Form.Dirty = False
'        Me.Refresh
'    End If
End Sub

Private Sub CurrMM_Enter()
    mstrMonth = "CurrMM"
End Sub

Private Sub CurrMM_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Dec_Enter()
    mstrMonth = "Dec"

End Sub

Private Sub Dec_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Feb_Enter()
    mstrMonth = "Feb"

End Sub

Private Sub Feb_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Jan_Enter()
    mstrMonth = "Jan"
End Sub

Private Sub Jan_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Jul_Enter()
    mstrMonth = "Jul"
End Sub

Private Sub Jul_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Jun_Enter()
    mstrMonth = "Jun"

End Sub

Private Sub Jun_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Mar_Enter()
    mstrMonth = "Mar"

End Sub

Private Sub Mar_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub May_Enter()
    mstrMonth = "May"

End Sub

Private Sub May_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Month1_Exit(Cancel As Integer)
    CheckDirty
End Sub

Private Sub Month10_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month11_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month12_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month2_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month3_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month4_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month5_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month6_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month7_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month8_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Month9_Exit(Cancel As Integer)
    CheckDirty

End Sub

Private Sub Nov_Enter()
    mstrMonth = "Nov"

End Sub

Private Sub Nov_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Oct_Enter()
    mstrMonth = "Oct"

End Sub

Private Sub Oct_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Sep_Enter()
    mstrMonth = "Sep"

End Sub

Private Sub Sep_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub Text28_Exit(Cancel As Integer)
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub

Private Sub ReCalcPercent()
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
End Sub
Private Sub cmdUp_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdUp_Click
    Select Case mstrMonth
    Case "Jan"
        Jan = Jan + 0.01
    Case "Feb"
        Feb = Feb + 0.01
    Case "Mar"
        Mar = Mar + 0.01
    Case "Apr"
        Apr = Apr + 0.01
    Case "May"
        May = May + 0.01
    Case "Jun"
        Jun = Jun + 0.01
    Case "Jul"
        Jul = Jul + 0.01
    Case "Aug"
        Aug = Aug + 0.01
    Case "Sep"
        Sep = Sep + 0.01
    Case "Oct"
        Oct = Oct + 0.01
    Case "Nov"
        Nov = Nov + 0.01
    Case "Dec"
        Dec = Dec + 0.01
    Case "CurrMM"
        CurrMM = CurrMM + 0.01

    End Select
    If Me.Dirty Then
        Me.Dirty = False
        Me.Refresh
    End If
Exit_cmdUp_Click:
    Exit Sub

Err_cmdUp_Click:
    MsgBox Err.Description
    Resume Exit_cmdUp_Click

End Sub
Private Sub cmdLower_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdLower_Click

    Select Case mstrMonth
    Case "Jan"
        Jan = Jan - 0.01
    Case "Feb"
        Feb = Feb - 0.01
    Case "Mar"
        Mar = Mar - 0.01
    Case "Apr"
        Apr = Apr - 0.01
    Case "May"
        May = May - 0.01
    Case "Jun"
        Jun = Jun - 0.01
    Case "Jul"
        Jul = Jul - 0.01
    Case "Aug"
        Aug = Aug - 0.01
    Case "Sep"
        Sep = Sep - 0.01
    Case "Oct"
        Oct = Oct - 0.01
    Case "Nov"
        Nov = Nov - 0.01
    Case "Dec"
        Dec = Dec - 0.01
    Case "CurrMM"
        CurrMM = CurrMM - 0.01

    End Select
    If Me.Dirty Then
        Me.Dirty = False
        Me.Refresh
    End If
Exit_cmdLower_Click:
    Exit Sub

Err_cmdLower_Click:
    MsgBox Err.Description
    Resume Exit_cmdLower_Click

End Sub
Private Sub Command32_Click()
If gcfHandleErrors Then On Error GoTo Err_Command32_Click

    Dim stDocName As String
    SetOptions False
    AssignDaysSupplyGroupIDs
Exit_Command32_Click:
    SetOptions True

    DoCmd.Close

    Exit Sub

Err_Command32_Click:
    MsgBox Err.Description
    Resume Exit_Command32_Click

End Sub
Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub

Private Sub CheckDirty()
    Dim nCurMM As Integer
    If Me.Form.Dirty Then
        Me.Form.Dirty = False
        Me.Refresh
    End If
    nCurMM = Month(Now)
    Select Case nCurMM
    Case 1 'January
        CurrMM = CurrentMonth
        Jan = Month12
        Feb = Month11
        Mar = Month10
        Apr = Month9
        May = Month8
        Jun = Month7
        Jul = Month6
        Aug = Month5
        Sep = Month4
        Oct = Month3
        Nov = Month2
        Dec = Month1
    Case 2 'FEB
        CurrMM = CurrentMonth
        Jan = Month1
        Feb = Month12
        Mar = Month11
        Apr = Month10
        May = Month9
        Jun = Month8
        Jul = Month7
        Aug = Month6
        Sep = Month5
        Oct = Month4
        Nov = Month3
        Dec = Month2
    Case 3 'MAR
        CurrMM = CurrentMonth
        Jan = Month2
        Feb = Month1
        Mar = Month12
        Apr = Month11
        May = Month10
        Jun = Month9
        Jul = Month8
        Aug = Month7
        Sep = Month6
        Oct = Month5
        Nov = Month4
        Dec = Month3
    Case 4 'APR
        CurrMM = CurrentMonth
        Jan = Month3
        Feb = Month2
        Mar = Month1
        Apr = Month12
        May = Month11
        Jun = Month10
        Jul = Month9
        Aug = Month8
        Sep = Month7
        Oct = Month6
        Nov = Month5
        Dec = Month4
    Case 5 'MAY
        CurrMM = CurrentMonth
        Jan = Month4
        Feb = Month3
        Mar = Month2
        Apr = Month1
        May = Month12
        Jun = Month11
        Jul = Month10
        Aug = Month9
        Sep = Month8
        Oct = Month7
        Nov = Month6
        Dec = Month5
    Case 6 'JUN
        CurrMM = CurrentMonth
        Jan = Month5
        Feb = Month4
        Mar = Month3
        Apr = Month2
        May = Month1
        Jun = Month12
        Jul = Month11
        Aug = Month10
        Sep = Month9
        Oct = Month8
        Nov = Month7
        Dec = Month6
    Case 7 'JUL
        CurrMM = CurrentMonth
        Jan = Month6
        Feb = Month5
        Mar = Month4
        Apr = Month3
        May = Month2
        Jun = Month1
        Jul = Month12
        Aug = Month11
        Sep = Month10
        Oct = Month9
        Nov = Month8
        Dec = Month7
    Case 8 'AUG
        CurrMM = CurrentMonth
        Jan = Month7
        Feb = Month6
        Mar = Month5
        Apr = Month4
        May = Month3
        Jun = Month2
        Jul = Month1
        Aug = Month12
        Sep = Month11
        Oct = Month10
        Nov = Month9
        Dec = Month8
    Case 9 'SEP
        CurrMM = CurrentMonth
        Jan = Month8
        Feb = Month7
        Mar = Month6
        Apr = Month5
        May = Month4
        Jun = Month3
        Jul = Month2
        Aug = Month1
        Sep = Month12
        Oct = Month11
        Nov = Month10
        Dec = Month9
    Case 10 'OCT
        CurrMM = CurrentMonth
        Jan = Month9
        Feb = Month8
        Mar = Month7
        Apr = Month6
        May = Month5
        Jun = Month4
        Jul = Month3
        Aug = Month2
        Sep = Month1
        Oct = Month12
        Nov = Month11
        Dec = Month10
    Case 11 'NOV
        CurrMM = CurrentMonth
        Jan = Month10
        Feb = Month9
        Mar = Month8
        Apr = Month7
        May = Month6
        Jun = Month5
        Jul = Month4
        Aug = Month3
        Sep = Month2
        Oct = Month1
        Nov = Month12
        Dec = Month11
    Case 12 'DEC
        CurrMM = CurrentMonth
        Jan = Month11
        Feb = Month10
        Mar = Month9
        Apr = Month8
        May = Month7
        Jun = Month6
        Jul = Month5
        Aug = Month4
        Sep = Month3
        Oct = Month2
        Nov = Month1
        Dec = Month12
    Case Else
        MsgBox "Month not found!" & Month(Now)
    End Select

    If Me.Dirty Then
        Me.Dirty = False
        Me.Refresh
    End If







End Sub
