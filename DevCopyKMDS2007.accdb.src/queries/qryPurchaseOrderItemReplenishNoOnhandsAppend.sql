PARAMETERS parVendorID Long; INSERT INTO PurchaseOrderItemDaysSupply (
  VendorID, Description, ItemId, SOQ,
  SPrice, Cost, Repack
)
SELECT
  Item.VendorID,
  Item.Description,
  Item.ID,
  Item.OrderTo AS SOQ,
  Item.Price,
  Item.Cost,
  0 AS Repack
FROM
  Item
  LEFT JOIN PurchaseOrderItemDaysSupply ON Item.ID = PurchaseOrderItemDaysSupply.ItemId
WHERE
  (
    (
      (Item.VendorID)= [parVendorID]
    )
    AND (
      (Item.OrderTo)> 0
    )
    AND (
      (
        PurchaseOrderItemDaysSupply.ItemId
      ) Is Null
    )
    AND (
      (Item.Onhand)<= 0
    )
  )
ORDER BY
  Item.SalesDepartmentID,
  Item.Description;
