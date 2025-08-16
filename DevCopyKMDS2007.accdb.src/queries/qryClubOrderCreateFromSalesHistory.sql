PARAMETERS parClubID Long;
SELECT
  InvoLine.Code,
  First(InvoLine.SetupDate) AS FirstOfSetupDate,
  Club.ID,
  Club.Name,
  Item.Description,
  Last(Item.Price) AS LastOfPrice,
  Last(Item.Cost) AS LastOfCost,
  Last(InvoLine.SetupDate) AS LastOfSetupDate,
  Sum(InvoLine.Quantity) AS SumOfQuantity
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
      (InvoLine.SalesDept)= [parClubID]
    )
  )
ORDER BY
  InvoLine.Code,
  First(InvoLine.SetupDate);
