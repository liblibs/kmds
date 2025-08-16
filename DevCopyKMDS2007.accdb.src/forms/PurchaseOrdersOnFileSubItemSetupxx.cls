Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
        UnitsPerOrder.BackColor = vbWhite

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    If ID > 1000 And ID < 99999 Then
        cmbVendorID = glnOLCCVendor
    Else
        If cmbVendorID = glnOLCCVendor Then
            MsgBox "This item can not be supplied by the OLCC!"
            Cancel = True
        End If
    End If
'    If Me.Onhand.OldValue <> Me.Onhand Then
'        AuditChange "Onhand", ID, Onhand.OldValue, Onhand.Value, Price
'    End If
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_BeforeUpdate"
    Resume ExitSub

End Sub

Private Sub Form_Current()
        txtAvailOnTheShelf = [Onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] - Nz(OnReserve.Column(1), 0)

End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog

    SetBackColor Description, gbAllowChangeDescription
    'SetBackColor Onhand, gbAllowChangeOnhand
    SetBackColor OrderAt, gbAllowChangeOrderAt
    SetBackColor OrderTo, gbAllowChangeOrderTo
    SetBackColor UnitsPerCase, gbAllowChangeOrderTo
    SetBackColor UnitsPerOrder, gbAllowChangeOrderTo

    SetBackColor Cost, gbAllowChangeCost
    SetBackColor Price, gbAllowChangePrice

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub SetBackColor(myTextBox As TextBox, Allowed As Boolean)
If gcfHandleErrors Then On Error GoTo ErrLog
    With myTextBox
        .Locked = Not Allowed
        .Enabled = Allowed

        If .Enabled Then
            .BackColor = vbWhite ' SetupDate.BackColor
        Else
            .BackColor = glnColorViewOnly
        End If
    End With
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "SetBackColor"
    Resume ExitSub

End Sub




Private Sub OrderTo_AfterUpdate()
    If OrderTo < UnitsPerOrder Then
        UnitsPerOrder = OrderTo - OrderAt
        UnitsPerOrder.BackColor = vbYellow
    End If
    If OrderTo > UnitsPerCase Then
        UnitsPerOrder = UnitsPerCase
        UnitsPerOrder.BackColor = vbYellow
    End If
End Sub

Private Sub txtAvailOnTheShelf_BeforeUpdate(Cancel As Integer)
    MsgBox "Onhand changes are not allowed here... Double click on Item Code to change onhands!"
    Cancel = True
    txtAvailOnTheShelf.Undo
End Sub
