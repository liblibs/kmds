SELECT
  Format(
    InvoLine.SetupDate, "mm/dd/yyyy"
  ) AS [Date],
  Sum(InvoLine.Price) AS SumOfPrice,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")<= 10,
      InvoLine.PRICE,
      0
    )
  ) AS [10 AM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 11,
      InvoLine.PRICE,
      0
    )
  ) AS [11 AM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 12,
      InvoLine.PRICE,
      0
    )
  ) AS [12 AM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 13,
      InvoLine.PRICE,
      0
    )
  ) AS [1 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 14,
      InvoLine.PRICE,
      0
    )
  ) AS [2 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 15,
      InvoLine.PRICE,
      0
    )
  ) AS [3 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 16,
      InvoLine.PRICE,
      0
    )
  ) AS [4 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 17,
      InvoLine.PRICE,
      0
    )
  ) AS [5 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 18,
      InvoLine.PRICE,
      0
    )
  ) AS [6 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 19,
      InvoLine.PRICE,
      0
    )
  ) AS [7 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")= 20,
      InvoLine.PRICE,
      0
    )
  ) AS [8 PM],
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "hh")>= 21,
      InvoLine.PRICE,
      0
    )
  ) AS [9 PM]
FROM
  (
    InvoLine
    INNER JOIN Item ON InvoLine.Code = Item.ID
  )
  INNER JOIN (
    SalesDepartment
    INNER JOIN SalesGroup ON SalesDepartment.SalesGroupID = SalesGroup.ID
  ) ON Item.SalesDepartmentID = SalesDepartment.ID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  )
GROUP BY
  Format(
    InvoLine.SetupDate, "mm/dd/yyyy"
  );
