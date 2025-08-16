SELECT
  Item.ID,
  Item.SizeID,
  Item.SalesDepartmentID,
  Item.VendorID,
  Item.SubCategoryID,
  vwItemsWithCurrentPricing.CurPrice,
  IIf([SPrice] <> 0, True, False) AS OnSale,
  Item.Cost,
  Item.Onhand,
  vwItemsWithCurrentPricing.OnShelf,
  Item.ChangeDate,
  Item.Description
FROM
  Item
  INNER JOIN vwItemsWithCurrentPricing ON Item.ID = vwItemsWithCurrentPricing.ID;
