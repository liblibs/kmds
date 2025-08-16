SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description,
  Sum(InvoLine.Price) AS SumOfPrice,
  Sum(
    IIf(
      [Tendercode] = 1, [InvoLine].[Price],
      0
    )
  ) AS Cash,
  Sum(
    IIf(
      [Tendercode] = 2, [InvoLine].[Price],
      0
    )
  ) AS Checks,
  Sum(
    IIf(
      [Tendercode] = 3, [InvoLine].[Price],
      0
    )
  ) AS CreditCard,
  Sum(
    IIf(
      [Tendercode] > 3, [InvoLine].[Price],
      0
    )
  ) AS Other,
  Sum(1) AS [Count]
FROM
  (
    qryInvoLineByFirstPaymentType
    INNER JOIN InvoLine ON (
      qryInvoLineByFirstPaymentType.CashSysID = InvoLine.CashSysID
    )
    AND (
      qryInvoLineByFirstPaymentType.InvoHeadID = InvoLine.InvoHeadID
    )
    AND (
      qryInvoLineByFirstPaymentType.SetupDate = InvoLine.SetupDate
    )
  )
  INNER JOIN (
    Item
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  ) ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.TransType)<> "T"
    )
  )
GROUP BY
  Item.SalesDepartmentID,
  SalesDepartment.Description
ORDER BY
  Item.SalesDepartmentID;
