SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description,
  Sum(InvoLine5.Price) AS SumOfPrice,
  Sum(
    IIf(
      [Tendercode] = 1, [InvoLine5].[Price],
      0
    )
  ) AS Cash,
  Sum(
    IIf(
      [Tendercode] = 2, [InvoLine5].[Price],
      0
    )
  ) AS Checks,
  Sum(
    IIf(
      [Tendercode] = 3, [InvoLine5].[Price],
      0
    )
  ) AS CreditCard,
  Sum(
    IIf(
      [Tendercode] > 3, [InvoLine5].[Price],
      0
    )
  ) AS Other,
  Sum(1) AS [Count]
FROM
  (
    (
      qryInvoLine5InvoheadByFirstPaymentType
      INNER JOIN InvoLine5 ON (
        qryInvoLine5InvoheadByFirstPaymentType.SetupDate = InvoLine5.SetupDate
      )
      AND (
        qryInvoLine5InvoheadByFirstPaymentType.CashSysID = InvoLine5.CashSysID
      )
      AND (
        qryInvoLine5InvoheadByFirstPaymentType.InvoHeadID = InvoLine5.InvoHeadID
      )
    )
    INNER JOIN Item ON InvoLine5.Code = Item.ID
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
WHERE
  (
    (
      (InvoLine5.TransType)<> "T"
    )
  )
GROUP BY
  Item.SalesDepartmentID,
  SalesDepartment.Description
ORDER BY
  Item.SalesDepartmentID;
