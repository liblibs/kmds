SELECT
  Tender.ID AS Department,
  Tender.Description AS DepartmentDesc,
  Sum(
    IIf([Quantity] = 1, 1, 0)
  ) AS TotalQ,
  Sum(InvoLine5.Price) AS TotalP,
  Sum(
    IIf(
      [CashsysID] = 1
      And [Quantity] = 1, 1,
      0
    )
  ) AS Q1,
  Sum(
    IIf(
      [CashsysID] = 1, InvoLine5.Price, 0
    )
  ) AS D1,
  Sum(
    IIf(
      [CashsysID] = 2
      And [Quantity] = 1, 1,
      0
    )
  ) AS Q2,
  Sum(
    IIf(
      [CashsysID] = 2, InvoLine5.Price, 0
    )
  ) AS D2,
  Sum(
    IIf(
      [CashsysID] = 3
      And [Quantity] = 1, 1,
      0
    )
  ) AS Q3,
  Sum(
    IIf(
      [CashsysID] = 3, InvoLine5.Price, 0
    )
  ) AS D3,
  Sum(
    IIf(
      [CashsysID] = 4
      And [Quantity] = 1, 1,
      0
    )
  ) AS Q4,
  Sum(
    IIf(
      [CashsysID] = 4, InvoLine5.Price, 0
    )
  ) AS D4
FROM
  InvoLine5
  INNER JOIN Tender ON InvoLine5.Code = Tender.ID
WHERE
  (
    (
      (InvoLine5.TransType)= "T"
    )
  )
GROUP BY
  Tender.ID,
  Tender.Description
ORDER BY
  Tender.ID;
