PARAMETERS parDate DateTime; INSERT INTO ItemMovementCurrentMonth (
  Code, Description, CurMQty, CurMDollars
)
SELECT
  InvoLine.Code,
  TranTypes.Description,
  Sum(InvoLine.Quantity) AS CurMQty,
  Sum(InvoLine.Price) AS CurMDollars
FROM
  InvoLine
  INNER JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
WHERE
  (
    (
      (InvoLine.SetupDate)>= [parDate]
    )
  )
GROUP BY
  InvoLine.Code,
  TranTypes.Description
ORDER BY
  InvoLine.Code;
