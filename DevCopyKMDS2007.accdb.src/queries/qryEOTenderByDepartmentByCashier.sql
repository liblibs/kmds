PARAMETERS parCashier Long;
SELECT
  Tender.ID AS Department,
  Tender.Description AS DepartmentDesc,
  Sum(InvoLine5.Quantity) AS TotalQ,
  Sum(InvoLine5.Price) AS TotalP,
  Sum(
    IIf(
      InvoHead5.CashsysID = 1, [Quantity],
      0
    )
  ) AS Q1,
  Sum(
    IIf(
      InvoHead5.CashsysID = 1, InvoLine5.Price,
      0
    )
  ) AS D1,
  Sum(
    IIf(
      InvoHead5.CashsysID = 2, [Quantity],
      0
    )
  ) AS Q2,
  Sum(
    IIf(
      InvoHead5.CashsysID = 2, InvoLine5.Price,
      0
    )
  ) AS D2,
  Sum(
    IIf(
      InvoHead5.CashsysID = 3, [Quantity],
      0
    )
  ) AS Q3,
  Sum(
    IIf(
      InvoHead5.CashsysID = 3, InvoLine5.Price,
      0
    )
  ) AS D3,
  Sum(
    IIf(
      InvoHead5.CashsysID = 4, [Quantity],
      0
    )
  ) AS Q4,
  Sum(
    IIf(
      InvoHead5.CashsysID = 4, InvoLine5.Price,
      0
    )
  ) AS D4,
  InvoLine5.TransType,
  InvoHead5.CashierID
FROM
  (
    InvoLine5
    INNER JOIN InvoHead5 ON (
      InvoLine5.InvoHeadID = InvoHead5.ID
    )
    AND (
      InvoLine5.CashSysID = InvoHead5.CashSysID
    )
  )
  INNER JOIN Tender ON InvoLine5.Code = Tender.ID
GROUP BY
  Tender.ID,
  Tender.Description,
  InvoLine5.TransType,
  InvoHead5.CashierID
HAVING
  (
    (
      (InvoHead5.CashierID)= [parCashier]
    )
  )
ORDER BY
  Tender.ID;
