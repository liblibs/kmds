SELECT
  InvoLine.Code,
  Min(InvoLine.SetupDate) AS MinOfSetupDate,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate,
  Club.ID,
  Club.Name,
  Item.Description,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Last(Item.Price) AS LastOfPrice,
  Sum(InvoLine.price) AS ExtPrice,
  Last(Item.Cost) AS LastOfCost
FROM
  (
    InvoLine
    INNER JOIN Item ON InvoLine.Code = Item.ID
  )
  INNER JOIN Club ON InvoLine.SalesDept = Club.ID
WHERE
  (
    (
      (InvoLine.TransType)= "C"
    )
  )
GROUP BY
  InvoLine.Code,
  Club.ID,
  Club.Name,
  Item.Description,
  InvoLine.SalesDept
HAVING
  (
    (
      (Club.ID)= [Forms]![ClubOrderCreate]![lstClubs]
    )
  )
ORDER BY
  InvoLine.Code,
  Min(InvoLine.SetupDate);
