SELECT
  Newprice.Code,
  Item.Description,
  CDate(
    Mid([startdate], 3, 2)& "/" & Right([startdate], 2)& "/" & Left([startDate], 2)
  ) AS SalePriceStartDate,
  CDate(
    Mid([Enddate], 3, 2)& "/" & Right([Enddate], 2)& "/" & Left([EndDate], 2)& " 23:59:59"
  ) AS SalePriceEndDate,
  IIf(
    IsNull(Item.Price),
    newprice.Price * 0.01,
    Item.price
  ) AS Old,
  [NewPrice] * 0.01 AS Sale,
  Item.Onhand,
  [Onhand] * Item.Price AS OldValue,
  [Onhand] *([NewPrice] * 0.01) AS NewValue,
  [onhand] *(
    ([NewPrice] * 0.01)- Item.Price
  ) AS Markup
FROM
  Newprice
  INNER JOIN Item ON Newprice.Code = Item.ID
WHERE
  (
    (
      (Newprice.StartDate)> " "
    )
    AND (
      (Newprice.EndDate)> "000000"
    )
  )
ORDER BY
  Newprice.Code;
