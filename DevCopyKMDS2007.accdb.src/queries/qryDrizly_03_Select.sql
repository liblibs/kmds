SELECT
  Settings.Value AS StoreLocationId,
  qryVWItemsWithCurrentPricing.ID AS ProductID,
  qryVWItemsWithCurrentPricing.Description,
  qryVWItemsWithCurrentPricing.CurPrice AS RetailPrice,
  qryVWItemsWithCurrentPricing.OnShelf AS QuantityOnHand,
  qryUPCItemMax.UPCCodeMax AS UPCCode,
  qryVWItemsWithCurrentPricing.PkgSize AS [Size],
  Left(
    [subcategory].[Description],
    Len([subcategory].[Description])-4
  ) AS OtherInfo,
  Item.QtyPack,
  Item.QtyPackPrice,
  Item.Deposit
FROM
  Sysfile,
  Settings,
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
      (
        qryVWItemsWithCurrentPricing.OnShelf
      )>-1
    )
    AND (
      (Settings.KeyName)= "DrizlyStoreID"
    )
    AND (
      (SubCategory.Description) Like "*ONL"
    )
  )
ORDER BY
  qryVWItemsWithCurrentPricing.Description;
