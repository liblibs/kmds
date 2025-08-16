PARAMETERS parClubID Long;
SELECT
  ClubOrderGuides.ClubID,
  ClubOrderGuides.OrderID,
  ClubOrderGuides.LineID,
  ClubOrderGuides.ItemID,
  ClubOrderGuides.Par,
  Item.Description,
  Item.Price,
  IIf(
    IsNull([SalePrice]),
    Item.Price,
    [SalePrice]
  ) AS SPrice,
  IIf(
    IsNull([SalePrice]),
    0,
    -1
  ) AS SP,
  Item.Cost,
  ClubOrderGuides.QuotedPrice,
  IIf([SalesDepartmentID] = 1,-1, 0) AS Liquor,
  IIf(
    IsNull(qryBuyDownGoingOnNow.SpcPrice),
    0,
    qryBuyDownGoingOnNow.SpcPrice
  ) AS Buydown,
  qryBuyDownGoingOnNow.ItemsCode
FROM
  (
    (
      Item
      LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
    )
    INNER JOIN ClubOrderGuides ON Item.ID = ClubOrderGuides.ItemID
  )
  LEFT JOIN qryBuyDownGoingOnNow ON Item.ID = qryBuyDownGoingOnNow.ItemsCode
WHERE
  (
    (
      (ClubOrderGuides.ClubID)= [parClubID]
    )
  )
ORDER BY
  ClubOrderGuides.LineID;
