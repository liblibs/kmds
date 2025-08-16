Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click

End Sub
'------------------------------------------------------------
' Command2_Click
'

'------------------------------------------------------------
Private Sub Command2_Click()
Dim comQDF As Object
Dim stDocName As String
On Error GoTo ErrLog

Set comQDF = CreateObject("ADODB.COMMAND")

    If MsgBox("Create Order from " & OrderDate & vbCrLf & "Amount: " & Format(Extended, "Currency") & "?", vbYesNo) = vbYes Then

        On Error Resume Next
        stDocName = "OLCCEPSNotOnFile"
        DoCmd.OpenReport stDocName, acViewPreview, , "OrderNumber=" & OrderNumber
        ArchiveReport stDocName
        DoCmd.Close acReport, stDocName
        DoCmd.OpenReport stDocName, acViewNormal, , "OrderNumber=" & OrderNumber
        DoCmd.Close acReport, stDocName

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryOLCCEPSForOrdersNotOnFileAppend"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("prmOrderNumber", adBigInt, adParamInput)
            .Parameters("prmOrderNumber") = OrderNumber
            .Parameters.Append .CreateParameter("prmOrderDate", adDate, adParamInput)
            .Parameters("prmOrderDate") = OrderDate
            .Execute
        End With
        AppendUPCItemsWithItemID
        'MsgBox "Create Purchase Order"
        CreatePurchaseOrderFromOLCCEPS 1, OrderNumber, OrderDate
        SetOptions False
        DoCmd.OpenQuery "qryOLCCEPS_Update_UnitsPerCase"
        SetOptions True

        On Error Resume Next
        DoCmd.Close 'acForm, Me
    End If
Exit_Command20_Click:
    Exit Sub

ErrLog:
    MsgBox "OLCC PACKING LIST: " & Err.Description
    Resume Exit_Command20_Click
End Sub


Private Sub Form_Current()
Dim comQDF As Object

Dim nOrderCount As Integer
Dim nPurchaseOrderID As Long
Dim nLineCount As Integer
Dim dtSetup As Date
Dim strMsg As String

Set comQDF = CreateObject("ADODB.COMMAND")

If gcfHandleErrors Then On Error GoTo ErrLog
    If Not IsNull(OrderNumber) Then
        'ArchiveReport "PurchaseOrderCreate"
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryOLCCEPSForOrders"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("prmOrderNumber", adInteger, adParamInput)
            .Parameters("prmOrderNumber") = OrderNumber
        End With
    End If
ExitSub:
    Exit Sub
ErrLog:
    Resume ExitSub
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim obj
    Me.Requery
    Set obj = Me.RecordsetClone
    If obj.RecordCount = 0 Then
        MsgBox "No records!"
        Cancel = True
    End If

End Sub
'------------------------------------------------------------
' cmdOLCCEPSPriceDiff_Click
'
'------------------------------------------------------------
Private Sub cmdOLCCEPSPriceDiff_Click()
On Error GoTo cmdOLCCEPSPriceDiff_Click_Err

    DoCmd.OpenForm "OLCCEPSPriceDiff", acNormal, , , , acDialog


cmdOLCCEPSPriceDiff_Click_Exit:
    Exit Sub

cmdOLCCEPSPriceDiff_Click_Err:
    MsgBox Error$
    Resume cmdOLCCEPSPriceDiff_Click_Exit

End Sub
