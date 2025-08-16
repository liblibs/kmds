SELECT
  Sum(
    IIf(
      [transType] = "S", InvoLine5.Price,
      0
    )
  ) AS [Counter],
  Sum(
    IIf(
      [transType] = "C", InvoLine5.Price,
      0
    )
  ) AS ClubsGross,
  Sum(
    IIf(
      [transType] = "D", InvoLine5.Price,
      0
    )
  ) AS ClubsDiscount,
  Sum(
    IIf(
      [transType] <> "S", InvoLine5.Price,
      0
    )
  ) AS ClubsNet,
  Sum(InvoLine5.Price) AS OLCCDeposit
FROM
  InvoLine5
  INNER JOIN Item ON InvoLine5.Code = Item.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
      Or (Item.SalesDepartmentID)= 2
    )
    AND (
      (InvoLine5.TransType)<> "T"
    )
  );
