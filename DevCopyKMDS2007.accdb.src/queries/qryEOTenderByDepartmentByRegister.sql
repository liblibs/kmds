SELECT
  Tender.ID AS Department,
  Tender.Description AS DepartmentDesc,
  InvoLine5.InvoHeadID,
  Sum(
    IIf([Quantity] =-1, 1, 0)
  ) AS TotalQ,
  Sum(0 - [Price]) AS TotalP,
  Sum(
    IIf(
      [CashsysID] = 1
      And [Quantity] = 1, 1,
      0
    )
  ) AS Q1,
  Sum(
    IIf(
      [CashsysID] = 1, 0 - InvoLine5.Price,
      0
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
      [CashsysID] = 2, 0 - InvoLine5.Price,
      0
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
      [CashsysID] = 3, 0 - InvoLine5.Price,
      0
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
      [CashsysID] = 4, 0 - InvoLine5.Price,
      0
    )
  ) AS D4,
  Sum(
    IIf(
      [CashsysID] = 5
      And [Quantity] = 1, 1,
      0
    )
  ) AS Q5,
  Sum(
    IIf(
      [CashsysID] = 5, 0 - [InvoLine5].[Price],
      0
    )
  ) AS D5
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
  Tender.Description,
  InvoLine5.InvoHeadID
ORDER BY
  Tender.ID;
