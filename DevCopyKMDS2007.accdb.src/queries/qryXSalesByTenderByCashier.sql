SELECT
  InvoLine.Code,
  Tender.Description,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Cashier.ID,
  Cashier.Name
FROM
  Cashier
  INNER JOIN (
    InvoLine
    INNER JOIN Tender ON InvoLine.Code = Tender.ID
  ) ON Cashier.ID = InvoLine.CashierID
WHERE
  (
    (
      (InvoLine.TransType)= 'T'
    )
  )
GROUP BY
  InvoLine.Code,
  Tender.Description,
  Cashier.ID,
  Cashier.Name
HAVING
  (
    (
      (Cashier.ID)= [parCashierID]
    )
  )
ORDER BY
  InvoLine.Code,
  Cashier.ID;
