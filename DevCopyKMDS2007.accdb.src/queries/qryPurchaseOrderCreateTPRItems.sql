PARAMETERS parVendorID Long;
SELECT
  Item.VendorID,
  Item.ID AS ItemID,
  Item.Description,
  Item.Price,
  qryItemsSalePriceNow.SalePrice,
  Item.Cost,
  Item.TPR
FROM
  (
    qryItemsSalePriceNow
    RIGHT JOIN Item ON qryItemsSalePriceNow.ItemID = Item.ID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (Item.VendorID)= [parVendorID]
    )
    AND (
      (Item.TPR)<> 0
    )
    AND (
      (SubCategory.Description)<> "OBSOLETE"
    )
  )
ORDER BY
  Item.ID;
