SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  Item.SalesDepartmentID AS POID,
  InvoLine.SalesGroup AS VendorID,
  Vendor.Name AS VendorName,
  Vendor.Contact,
  Vendor.Address1,
  Vendor.Address2,
  [City] & ", " & [St] & "  " & [Zip] AS CitySTZ,
  Vendor.Phone1,
  Vendor.PoAtCost,
  Vendor.PrintUpcOnPo,
  InvoLine.ID AS LineID,
  InvoLine.Code AS ItemID,
  Item.Description,
  InvoLine.Quantity AS QuantityShipped,
  InvoLine.Cost / InvoLine.quantity AS Cost,
  InvoLine.price / InvoLine.quantity AS Price,
  Size.Description AS SizeDesc,
  [Quantity] \[UnitsPerCase] AS Cases,
  Item.UnitsPerCase,
  InvoLine.Cost AS ExtCost,
  InvoLine.Price AS ExtPrice,
  Item.StockNumber,
  Vendor.Phone2 AS FAX,
  qryUPCItemMax.UPCCodeMax
FROM
  (
    (
      (
        (
          (
            InvoLine
            LEFT JOIN Item ON InvoLine.Code = Item.ID
          )
          LEFT JOIN Cashier ON InvoLine.CashierID = Cashier.ID
        )
        LEFT JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
      )
      INNER JOIN Vendor ON InvoLine.SalesGroup = Vendor.ID
    )
    LEFT JOIN [Size] ON Item.SizeID = Size.ID
  )
  LEFT JOIN qryUPCItemMax ON Item.ID = qryUPCItemMax.ItemID
WHERE
  (
    (
      (InvoLine.InvoHeadID)= Forms!Vendors!VendorsPurchaseOrdersSub.Form!InvoHeadID
    )
    And (
      (InvoLine.CashSysID)= Forms!Vendors!VendorsPurchaseOrdersSub.Form!CashSysID
    )
    And (
      (InvoLine.SetupDate)= Forms!Vendors!VendorsPurchaseOrdersSub.Form!Date
    )
    And (
      (InvoLine.Quantity)<> 0
    )
  );
