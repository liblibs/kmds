INSERT INTO ItemsSoldToday (
  Code, SoldToday, DollarsSoldToday
)
SELECT
  Item.ID,
  Sum(
    IIf(
      [TransType] = "x", 0 - [Quantity], [Quantity]
    )
  ) AS Qty,
  Sum(InvoLine5.Price) AS SumOfPrice
FROM
  InvoLine5
  INNER JOIN Item ON InvoLine5.Code = Item.ID
GROUP BY
  Item.ID;
