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
    'If ID > 1000 And ID < 99999 Then
    If ID > 99900000000# Then
           '99900014075
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
    If Me.Price.OldValue <> Me.Price Then
        'DoCmd.OpenForm "AdjustReasons", acNormal, , , acFormEdit, acDialog

        AuditPriceChange "Price", ID, Price.OldValue, Price.Value, Onhand
    End If


ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_BeforeUpdate"
    Resume ExitSub

End Sub

Private Sub Form_Current()
        txtDaysSupply.BackColor = vbWhite
        If gcfHandleErrors Then On Error Resume Next
        txtAvailOnTheShelf = [Onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] - Nz(OnReserve.Column(1), 0)
        If AverageDailySales <> 0 And txtAvailOnTheShelf <> 0 Then

            txtDaysSupply = CInt(txtAvailOnTheShelf / AverageDailySales)
            If txtDaysSupply < 10 Then
                txtDaysSupply.BackColor = 8421631 'pale vbRed
            Else
                If txtDaysSupply < 20 Then
                    txtDaysSupply.BackColor = 8454143 'pale vbYellow
                End If
            End If

        Else
            txtDaysSupply = 0
        End If
        txtAvailOnTheShelf.BackColor = txtDaysSupply.BackColor

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
