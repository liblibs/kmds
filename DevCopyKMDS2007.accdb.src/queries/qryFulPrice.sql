SELECT
  Item.ID,
  Item.Description,
  Item.Onhand,
  Item.Price,
  Format(
    ([fulprice].[Price] * 0.01),
    "Currency"
  ) AS WebPrice,
  Left([MInOfCode], 4) AS [Year],
  Mid([MinOfCode], 5, 2) AS [Month],
  Fulprice.Price,
  Fulprice.Code
FROM
  qryFulPriceMinCode,
  Fulprice
  INNER JOIN Item ON Fulprice.Code = Item.ID
WHERE
  (
    (
      (Item.Price)<>([fulprice].[Price] * 0.01)
    )
  )
ORDER BY
  Item.ID;
