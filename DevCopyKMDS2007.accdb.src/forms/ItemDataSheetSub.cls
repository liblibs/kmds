Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Dim gbRunMailChanges As Boolean


Private Sub Form_AfterUpdate()

Dim nsaveID As Double '07/17/2021
Dim rs As Object

    If Not IsNull(PriceGroupID) And PriceGroupID > 0 Then
        PriceGroupUpdate PriceGroupID, Price, QtyPack, QtyPackPrice, Cost
        nsaveID = ID
        'to reflect the new prices on the 'other' items in the price group
        Me.Requery
        Set rs = Me.Recordset.Clone
        rs.FindFirst "[ID] = " & Str(nsaveID)
        Me.Bookmark = rs.Bookmark
    End If
    gbRunMailChanges = True

End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    'MsgBox "Applying: " & Me.Filter
    [Forms]![ItemDataSheetSelect]!txtFilter = Me.Filter
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim strSQL As String
    If Not IsNull(PriceGroupID) And PriceGroupID > 0 And Price.Value <> Price.OldValue Then
        strSQL = "This will change every item in this Price Group (" & PriceGroupID.Column(1) & ")" & vbCrLf
        strSQL = strSQL & "to " & Format(Price, "Currency") & vbCrLf & vbCrLf
        strSQL = strSQL & "Are you sure?"
        If MsgBox(strSQL, vbYesNo) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    AuditChange "Onhand", ID, Onhand.OldValue, Onhand.Value, Price

    ChangeDate = Now
    gbRunMailChanges = True
End Sub

Private Sub Form_Current()

 ' Declare and set a variable to store the WHERE
 ' clause that describes the records you want to
 ' display.
 Dim strCond As String
 If IsNull(ID) Then
    strCond = "Code = -1"
 Else
     strCond = "Code = " & Str(ID)
 End If

 ' Use the IsLoaded function from the Northwind
 ' sample database to check whether the Products
 ' form is open, then set the properties.
 'If IsLoaded("PurchaseOrdersOnFileSubItemMovement") Then
 'Forms![ItemDataSheet]![qryItemMovementByMonth].Form![Code]
    Forms![itemdatasheet]![qryItemMovementByMonth].Form.Filter = strCond
    Forms![itemdatasheet]![qryItemMovementByMonth].Form.FilterOn = True
 'End If

End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    'MsgBox "Form_filter: " & Me.Filter

End Sub

Private Sub Form_Load()
If gcfHandleErrors Then On Error GoTo ErrLog
    ID.Visible = True
    ID.ColumnHidden = False
    ID.SetFocus 'Me.SetFocus
    SetBackColor Price, gbAllowChangePrice
    SetBackColor Onhand, gbAllowChangeOnhand
    SetBackColor Description, gbAllowChangeDescription

    SetBackColor Cost, gbAllowChangeCost
    SetBackColor OrderTo, gbAllowChangeOrderTo
    SetBackColor OrderAt, gbAllowChangeOrderAt

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub

Private Sub SetBackColor(myTextBox As TextBox, Allowed As Boolean)
    With myTextBox
        .Locked = Not Allowed
        .Enabled = Allowed

        If .Enabled Then
            .BackColor = SetupDate.BackColor
        Else
            .BackColor = ID.BackColor
        End If
    End With
End Sub
