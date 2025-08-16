PARAMETERS parInvoiceNumber Long,
parCashSysID Long,
parCashierID Long,
parShiftID Long,
parPurchaseOrderID Long; INSERT INTO InvoLine (
  InvoHeadID, CashSysID, ID, CashierID,
  ShiftID, Code, SalesDept, Quantity,
  Price, InputCode, SalesGroup, TransType,
  SetupDate
)
SELECT
  [parInvoiceNumber] AS InvoHeadID,
  [parCashSysID] AS CashSysID,
  PurchaseOrderLine.LineID AS ID,
  [parCashierID] AS CashierID,
  [parShiftID] AS ShiftID,
  PurchaseOrderLine.ItemID AS Code,
  Item.SalesDepartmentID AS SalesDept,
  PurchaseOrderLine.QuantityShipped AS Quantity,
  [QuantityShipped] * Item.Price AS XPrice,
  Left(
    CStr([PurchaseOrderID])& " " & Vendor.Name,
    15
  ) AS InputCode,
  Item.Onhand AS SalesGroup,
  "P" AS TransType,
  Now() AS SetupDate
FROM
  (
    PurchaseOrderLine
    INNER JOIN Item ON PurchaseOrderLine.ItemID = Item.ID
  )
  INNER JOIN Vendor ON PurchaseOrderLine.VendorID = Vendor.ID
WHERE
  (
    (
      (
        PurchaseOrderLine.QuantityShipped
      )<> 0
    )
    AND (
      (
        PurchaseOrderLine.PurchaseOrderID
      )= [parPurchaseOrderID]
    )
  );
