SELECT
  Newprice.Code,
  Item.Description,
  Item.Price AS Old,
  CCur(
    Format(Newprice.NewPrice * 0.01, "Fixed")
  ) AS New,
  Item.Onhand,
  [onhand] * Item.Price AS OldValue,
  CCur(
    Format(
      [onhand] * Newprice.NewPrice * 0.01,
      "Fixed"
    )
  ) AS NewValue,
  CDate(
    Mid([startdate], 3, 2)& "/" & Right([startdate], 2)& "/" & Left([startDate], 2)
  ) AS NewPriceStartDate,
  IIf(
    (
      (Newprice.NewPrice * 0.01)- Item.Price
    )> 0,
    (
      (Newprice.NewPrice * 0.01)- Item.Price
    )* [onhand],
    0
  ) AS Markup,
  IIf(
    (
      (Newprice.NewPrice * 0.01)- Item.Price
    )<= 0,
    (
      (Newprice.NewPrice * 0.01)- Item.Price
    )* [onhand],
    0
  ) AS MarkDown
FROM
  Newprice
  INNER JOIN Item ON Newprice.Code = Item.ID
WHERE
  (
    (
      (
        CCur(
          Format(Newprice.NewPrice * 0.01, "Fixed")
        )
      )<> item.Price
    )
    And (
      (Newprice.StartDate)> " "
    )
    And (
      (Newprice.EndDate)= "000000"
    )
  )
ORDER BY
  Newprice.Code;
