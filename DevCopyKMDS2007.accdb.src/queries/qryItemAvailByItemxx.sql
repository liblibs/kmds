PARAMETERS parItemID Long;
SELECT
  CStr(Item.ID) AS Code,
  CStr(Item.Description) AS Name,
  [Onhand] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] + [TransferredToday] - [Reserved] AS Avail,
  CCur(
    Format(
      IIf(
        IsNull([SalePrice]),
        Item.Price,
        [SalePrice]
      ),
      '#.00'
    )
  ) AS Price,
  Format(
    IIf(
      IsNull(
        [qryBuyDownGoingOnNow].[SpcPrice]
      ),
      0,
      [qryBuyDownGoingOnNow].[SpcPrice]
    ),
    '#.00'
  ) AS Buydown,
  Item.SalesDepartmentID,
  Item.Price AS RegularPrice,
  IIf(
    IsNull([OnReserve]),
    0,
    [OnReserve]
  ) AS Reserved,
  [Price] - [Buydown] AS CurPrice
FROM
  (
    (
      Item
      LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
    )
    LEFT JOIN qryBuyDownGoingOnNow ON Item.ID = qryBuyDownGoingOnNow.ItemsCode
  )
  LEFT JOIN qryItemOnReserve ON Item.ID = qryItemOnReserve.ItemID
WHERE
  (
    (
      (
        CStr([Item].[ID])
      )= [parItemID]
    )
  )
ORDER BY
  Item.ID;
