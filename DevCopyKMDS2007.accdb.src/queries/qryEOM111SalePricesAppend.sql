INSERT INTO ItemSalePrice (
  ItemID, StartDate, EndDate, OldPrice,
  SalePrice
)
SELECT
  Newprice.Code,
  CDate(
    Mid([startdate], 3, 2)& "/" & Right([startdate], 2)& "/" & Left([startDate], 2)
  ) AS NewPriceStartDate,
  CDate(
    Mid([Enddate], 3, 2)& "/" & Right([Enddate], 2)& "/" & Left([EndDate], 2)& " 23:59:59"
  ) AS NewPriceEndDate,
  IIf(
    IsNull(Item.Price),
    newprice.Price * 0.01,
    Item.price
  ) AS Old,
  [NewPrice] * 0.01 AS New
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
  );
