Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

''------------------------------------------------------------
'' Command2_Click
''
''------------------------------------------------------------
'Private Sub Command2_Click()
'On Error GoTo Command2_Click_Err
'    SetOptions False
'    DoCmd.OpenQuery "qryPricingWorkTable_3_UpdateNewPrice", acViewNormal, acEdit
'    DoCmd.OpenQuery "qryPricingWorkTable_4_UpdateNewGP", acViewNormal, acEdit
'    [PricingWorkTable_subform].[Form].Requery
'Command2_Click_Exit:
'    Exit Sub
'
'Command2_Click_Err:
'    MsgBox Error$
'    Resume Command2_Click_Exit
'End Sub


'------------------------------------------------------------
' cmdCalcPercentIncrease_Click
'
'------------------------------------------------------------
Private Sub cmdCalcPercentIncrease_Click()
On Error GoTo cmdCalcPercentIncrease_Click_Err
    frmPennies.Value = 1

    SetOptions False
    DoCmd.OpenQuery "qryPricingWorkTable_3_UpdateNewPrice", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPricingWorkTable_4_UpdateNewGP", acViewNormal, acEdit
    [PricingWorkTable_subform].[Form].Requery
    If txtPercent > 0 Then
        cmdDown.Visible = True
    Else
        cmdUp.SetFocus
        cmdDown.Visible = False
    End If



cmdCalcPercentIncrease_Click_Exit:
    SetOptions True
    Exit Sub

cmdCalcPercentIncrease_Click_Err:
    MsgBox Error$
    Resume cmdCalcPercentIncrease_Click_Exit

End Sub


Private Sub cmdDown_Click()
On Error GoTo cmdDown_Click_Err

    If txtPercent > 0 Then
        txtPercent = txtPercent - 1
        cmdCalcPercentIncrease_Click

    End If

cmdDown_Click_Exit:
    Exit Sub

cmdDown_Click_Err:
    MsgBox Error$
    Resume cmdDown_Click_Exit
End Sub

Private Sub cmdfrmClose_Click()
If gcfHandleErrors Then On Error GoTo ErrLog

    DoCmd.Close

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdfrmClose_Click"
    Resume ExitSub
End Sub

'------------------------------------------------------------
' cmdUpdatePrices_Click
'
'------------------------------------------------------------
Private Sub cmdUpdatePrices_Click()
If gcfHandleErrors Then On Error GoTo cmdUpdatePrices_Click_Err
    If MsgBox("CHANGE PRICES OF ALL ITEMS IN YELLOW" & vbCrLf & "Are you sure?", vbCritical + vbYesNo, "Prices will change!!!") <> vbYes Then
        Exit Sub
    End If

Dim strSQL As String
Dim nVendorID As Long
'Dim rsCashsys As New ADODB.Recordset
Dim rsCashsys As Object
Set rsCashsys = CreateObject("ADODB.Recordset")

'Dim rsInvoLine As New ADODB.Recordset
Dim rsInvoLine As Object
Set rsInvoLine = CreateObject("ADODB.Recordset")

'Dim rsItemMovement As New ADODB.Recordset
Dim rsItemMovement As Object
Set rsItemMovement = CreateObject("ADODB.Recordset")

'Dim rsSalePrice As New ADODB.Recordset
Dim rsSalePrice As Object
Set rsSalePrice = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Dim nLineCount As Integer


nLineCount = 1
Set comQDF = CreateObject("ADODB.COMMAND")

    Dim nLastInvoiceNo As Long
    Dim nShiftCounter As Long
    Dim CurPrice As Currency
    Dim curNewPrice As Currency

    SetOptions False
    '1. get the next invoice #
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rsCashsys.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rsCashsys.BOF And rsCashsys.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
        Exit Sub
    End If

    rsCashsys!LastInvoiceNo = rsCashsys!LastInvoiceNo + 1
    nLastInvoiceNo = rsCashsys!LastInvoiceNo + 1000000
    nShiftCounter = rsCashsys!ShiftCounter
    rsCashsys.Update
    rsCashsys.Close
    Set rsCashsys = Nothing

    On Error GoTo cmdUpdatePrices_Click_Err

    rsInvoLine.Open "InvoLine", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable


    Dim rs As Object
    Set rs = [PricingWorkTable_subform].[Form].RecordsetClone
    With rs
    Do While Not .EOF
        If rs!Price <> rs!NewPrice Then 'skip over items not changed
            nVendorID = rs!VendorID
            strSQL = "UPDATE Item SET Item.Price = " & rs!NewPrice & ", "
            strSQL = strSQL & "Item.ChangeDate = Now() "
            strSQL = strSQL & "WHERE (((Item.ID)=" & rs!ID & "))"
            'Debug.Print strsql
            DoCmd.RunSQL strSQL
            With rsInvoLine
                .AddNew
                !InvoHeadID = nLastInvoiceNo
                !ID = nLineCount 'nLineCount
                !CashSysID = glnCashsysID
                !CashierID = glnCashierID
                !ShiftID = nShiftCounter
                !Code = rs!ID
                CurPrice = rs!Price
                curNewPrice = rs!NewPrice
                !InputCode = Left(CStr(CurPrice & " to " & curNewPrice), 15) 'gstrAdjustReason
                !SalesDept = 0 'Onhand 'rsItem!SalesDepartmentID
                !Quantity = 0 'Onhand '!QuantityShipped
                !Price = curNewPrice
                !SalesGroup = (rs!Price * 100)
                !TransType = "R" 'Retail price change
                !SetupDate = Now
                .Update
                nLineCount = nLineCount + 1
            End With
        End If
        .MoveNext
    Loop
    End With
    CreatePurchaseOrderFromPricingTool nVendorID
    MsgBox "Purchase order created.. Print new PRICE LABELS in the C-1 Purchase order screen", vbApplicationModal
cmdUpdatePrices_Click_Exit:
    SetOptions True

    DoCmd.Close , ""

    Exit Sub

cmdUpdatePrices_Click_Err:
    MsgBox Error$
    Resume '    cmdUpdatePrices_Click_Exit

End Sub


Private Sub frmPennies_AfterUpdate()
On Error GoTo cmdAdjustPennies_Click_Err

SetOptions False
Select Case frmPennies
Case 1
    DoCmd.OpenQuery "qryPricingWorkTable_3_UpdateNewPrice", acViewNormal, acEdit

Case 2
    DoCmd.OpenQuery "qryPricingWorkTable_3_UpdateNewPrice", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case1", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case2", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case3", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case4", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9Case6", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9Case7", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9Case8", acViewNormal, acEdit
Case 3
    DoCmd.OpenQuery "qryPricingWorkTable_3_UpdateNewPrice", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case1", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case2", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case3", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo5Case4", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9Case6", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9Case7", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9Case8", acViewNormal, acEdit
    DoCmd.OpenQuery "qryPriceTo9CaseChange9to0", acViewNormal, acEdit
Case Else
    MsgBox "unhandled case Adjust Pennies"
End Select



cmdAdjustPennies_Click_Exit:
    DoCmd.OpenQuery "qryPricingWorkTable_4_UpdateNewGP", acViewNormal, acEdit
    [PricingWorkTable_subform].[Form].Requery
    SetOptions True
Exit Sub

cmdAdjustPennies_Click_Err:
    MsgBox Error$
    Resume cmdAdjustPennies_Click_Exit

End Sub
'------------------------------------------------------------
' cmdUp_Click
'
'------------------------------------------------------------
Private Sub cmdUp_Click()
On Error GoTo cmdUp_Click_Err

    If txtPercent < 99 Then
        txtPercent = txtPercent + 1
        cmdCalcPercentIncrease_Click
    End If

cmdUp_Click_Exit:
    Exit Sub

cmdUp_Click_Err:
    MsgBox Error$
    Resume cmdUp_Click_Exit

End Sub



Private Sub txtPercent_AfterUpdate()
    cmdCalcPercentIncrease_Click
End Sub
