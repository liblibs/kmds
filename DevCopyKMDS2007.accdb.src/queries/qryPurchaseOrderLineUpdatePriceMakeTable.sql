SELECT
  Item.ID,
  Item.Description,
  Item.Onhand,
  Item.Cost,
  IIf(
    IsNull([SalePrice]),
    [Price],
    [SalePrice]
  ) AS Sprice,
  IIf(
    IsNull([SalePrice]),
    " ",
    "*"
  ) AS Star,
  Item.VendorID AS ItemVendorID,
  Item.UnitsPerCase,
  Size.Description AS PkgSize,
  Item.SalesDepartmentID,
  Size.Oz INTO PurchaseOrderLineUpdatePrice
FROM
  (
    Item
    LEFT JOIN [Size] ON Item.SizeID = Size.ID
  )
  LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID;
