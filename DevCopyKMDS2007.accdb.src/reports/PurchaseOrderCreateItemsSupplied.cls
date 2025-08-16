Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)

    lblNoData.Visible = True
End Sub

Private Sub Report_Open(Cancel As Integer)




Dim strSQL As String
strSQL = "SELECT qryPurchaseOrderItemInfo.VendorID, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Name, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.ItemID, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Description, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.SizeID, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Onhand, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Ordered, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Reserved, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Avail, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.OrderAt, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.OrderTo, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.SOQRaw, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.OrderMultiple, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.SOQ, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.UnitsPerCase, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.SPrice, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.SPriceStar, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.ExtPrice, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.Cost, "
strSQL = strSQL & "qryPurchaseOrderItemInfo.ExtCost "
strSQL = strSQL & "FROM qryPurchaseOrderItemInfo "
strSQL = strSQL & "WHERE (((qryPurchaseOrderItemInfo.VendorID)= " & Forms("PurchaseOrderCreate").lstVendors & ")) "
'strSQL = strSQL & "AND ((qryPurchaseOrderItemInfo.Onhand)<=[OrderAt]) "
'strSQL = strSQL & "AND ((qryPurchaseOrderItemInfo.OrderTo)>0) "
'strSQL = strSQL & "AND ((qryPurchaseOrderItemInfo.SOQ)>0))"

Me.RecordSource = strSQL
End Sub
