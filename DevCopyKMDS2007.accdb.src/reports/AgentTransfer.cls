Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)
    If Status = "I" Then
        lblShipper.Caption = "Shipped From"
    End If
    If Status = "O" Then
        lblShipper.Caption = "Shipped To"
    End If
End Sub

Private Sub Report_Open(Cancel As Integer)
'Dim qdftemp As DAO.QueryDef
'    Set qdftemp = CurrentDb().QueryDefs("qryPurchaseOrder")
'    qdftemp.Parameters("parPurchaseOrderID") = Forms!PurchaseOrdersOnFile.PurchaseOrderID
'    Me.RecordSource = qdftemp.SQL ' .OpenRecordset

   ' If gcfHandleErrors Then On Error Resume Next
    Me.OrderByOn = True

    Me.OrderBy = Forms("AgentsTransfersChoices").TransferSort
                 'Forms![AgentsTransfersChoices]![TransferSor
    If Not gcfHandleErrors Then Debug.Print Forms("AgentsTransfersChoices").TransferSort
End Sub

Private Sub Report_Page()
    If Page = txtPages Then
        Me.PageFooterSection.Visible = False
    Else
        Me.PageFooterSection.Visible = True
    End If
End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Status = "I" Then
        lblTransfer.Caption = "Transfer IN"
    End If
    If Status = "O" Then
        lblTransfer.Caption = "Transfer OUT"
    End If
    If Forms![AgentsTransfersChoices]![optChoices] > 1 Then
        txtPosted = "Posted " & Format(Now, "mm/dd/yy hh:mm")
        txtPosted.Visible = True

    Else
        txtPosted.Visible = False
    End If

ExitSub:
    Exit Sub
ErrLog:
    MsgBox "error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub
