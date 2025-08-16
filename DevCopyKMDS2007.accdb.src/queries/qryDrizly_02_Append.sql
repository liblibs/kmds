INSERT INTO Drizly (
  StoreLocationId, ProductID, Description,
  RetailPrice, QuantityOnHand, UPCCode,
  [Size], OtherInfo, PackQty, PackRetailPrice
)
SELECT
  "OLCC " & [sysfile].[ID] AS StoreLocationId,
  qryVWItemsWithCurrentPricing.ID AS ProductID,
  qryVWItemsWithCurrentPricing.Description,
  qryVWItemsWithCurrentPricing.CurPrice AS RetailPrice,
  qryVWItemsWithCurrentPricing.OnShelf AS QuantityOnHand,
  qryUPCItemMax.UPCCodeMax AS UPCCode,
  qryVWItemsWithCurrentPricing.PkgSize AS [Size],
  SubCategory.Description,
  Item.QtyPack,
  Item.QtyPackPrice
FROM
  Sysfile,
  SubCategory
  INNER JOIN (
    (
      qryVWItemsWithCurrentPricing
      INNER JOIN qryUPCItemMax ON qryVWItemsWithCurrentPricing.ID = qryUPCItemMax.ItemID
    )
    INNER JOIN Item ON qryVWItemsWithCurrentPricing.ID = Item.ID
  ) ON SubCategory.ID = Item.SubCategoryID
WHERE
  (
    (
      (SubCategory.Description) Like "*ONL"
    )
  )
ORDER BY
  qryVWItemsWithCurrentPricing.Description;
