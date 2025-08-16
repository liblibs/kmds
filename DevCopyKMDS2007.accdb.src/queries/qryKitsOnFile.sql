SELECT
  Kit.ID,
  Kit.ItemID,
  Item.Description,
  Kit.Price,
  IIf(
    IsNull([Retail]),
    0,
    [retail]
  ) AS RetailPrice,
  [RetailPrice] - [Kit].[Price] AS Discount
FROM
  (
    Kit
    LEFT JOIN qryKitItemsTotalRetail ON Kit.ID = qryKitItemsTotalRetail.KitID
  )
  INNER JOIN Item ON Kit.ItemID = Item.ID;
