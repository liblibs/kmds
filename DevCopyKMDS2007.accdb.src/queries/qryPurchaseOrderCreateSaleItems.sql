PARAMETERS parVendorID Long;
SELECT
  Item.VendorID,
  qryItemsSalePriceNow.ItemID,
  Item.Description,
  Item.Price,
  qryItemsSalePriceNow.SalePrice,
  Item.Cost,
  SubCategory.Description
FROM
  (
    qryItemsSalePriceNow
    INNER JOIN Item ON qryItemsSalePriceNow.ItemID = Item.ID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (Item.VendorID)= [parVendorID]
    )
    AND (
      (SubCategory.Description)<> "OBSOLETE"
    )
  )
ORDER BY
  qryItemsSalePriceNow.ItemID;
