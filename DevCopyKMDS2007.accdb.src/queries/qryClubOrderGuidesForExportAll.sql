SELECT
  ClubOrderGuides.ClubID,
  Club.Name AS ClubName,
  ClubOrderGuides.LineID,
  ClubOrderGuides.ItemID,
  qryPurchaseOrderItemsWithTodaysPrice.Description,
  ClubOrderGuides.Par,
  qryPurchaseOrderItemsWithTodaysPrice.Sprice AS Price,
  qryPurchaseOrderItemsWithTodaysPrice.Star AS OnSale,
  Format(
    Now(),
    "mm/dd/yyyy"
  ) AS AsOf,
  vwItemsWithCurrentPricing.CategoryDescr AS Category
FROM
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
      (ClubOrderGuides.ClubID)= [parClubID]
    )
  )
ORDER BY
  ClubOrderGuides.LineID;
