SELECT
  ClubOrderGuides.LineID AS Line,
  0 AS Qty,
  ClubOrderGuides.ItemID,
  qryPurchaseOrderItemsWithTodaysPrice.Description,
  ClubOrderGuides.Par,
  qryPurchaseOrderItemsWithTodaysPrice.Sprice AS Price,
  [Qty] * [Price] AS Liquor,
  0 AS NonLiquor,
  "'" & CStr(
    Format(
      Now(),
      "mmm d"", ""yyyy"
    )
  ) AS AsOf,
  vwItemsWithCurrentPricing.CategoryDescr AS Category,
  qryPurchaseOrderItemsWithTodaysPrice.PkgSize,
  ([Sprice] / [Oz])* 1.25 AS PricePer125,
  ClubOrderGuides.ClubID,
  Club.Name AS ClubName,
  Sysfile.Name AS StoreName
FROM
  Sysfile,
  (
    (
      ClubOrderGuides
      INNER JOIN qryPurchaseOrderItemsWithTodaysPrice ON ClubOrderGuides.ItemID = qryPurchaseOrderItemsWithTodaysPrice.ID
    )
    INNER JOIN Club ON ClubOrderGuides.ClubID = Club.ID
  )
  INNER JOIN vwItemsWithCurrentPricing ON ClubOrderGuides.ItemID = vwItemsWithCurrentPricing.ID
WHERE
  (
    (
      (ClubOrderGuides.ClubID)= [Forms]![Clubs]![ID]
    )
  )
ORDER BY
  ClubOrderGuides.LineID;
