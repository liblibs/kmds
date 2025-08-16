SELECT
  Format(
    InvoLine.SetupDate, "mm/dd/yyyy"
  ) AS Today,
  Sum(
    IIf(
      [transType] = "P", InvoLine.Price,
      0
    )
  ) AS Purchases,
  Sum(
    IIf(
      [transType] = "X"
      And InvoLine.Price < 0,
      InvoLine.Price, 0
    )
  ) AS TransfersOUT,
  Sum(
    IIf(
      [transType] = "X"
      And InvoLine.Price > 0,
      InvoLine.Price, 0
    )
  ) AS TransfersIN,
  Sum(
    IIf(
      [transType] = "A", InvoLine.Price,
      0
    )
  ) AS Adjustments,
  Sum(
    IIf(
      [transType] = "S", 0 - InvoLine.Price,
      0
    )
  ) AS [Counter],
  Sum(
    IIf(
      [transType] = "C", 0 - InvoLine.Price,
      0
    )
  ) AS Clubs
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (InvoLine.SetupDate)> Date()
    )
  )
GROUP BY
  Format(
    InvoLine.SetupDate, "mm/dd/yyyy"
  );
