SELECT
  Format(InvoLine.SetupDate, "ww") AS Week,
  Min(
    Format(
      InvoLine.SetupDate, "mm/dd/yyyy"
    )
  ) AS [Date],
  Sum(InvoLine.Price) AS SumOfPrice,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Sun",
      InvoLine.PRICE,
      0
    )
  ) AS Sun,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Mon",
      InvoLine.PRICE,
      0
    )
  ) AS Mon,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Tue",
      InvoLine.PRICE,
      0
    )
  ) AS Tue,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Wed",
      InvoLine.PRICE,
      0
    )
  ) AS Wed,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Thu",
      InvoLine.PRICE,
      0
    )
  ) AS Thu,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Fri",
      InvoLine.PRICE,
      0
    )
  ) AS Fri,
  Sum(
    IIf(
      Format(InvoLine.SetupDate, "ddd")= "Sat",
      InvoLine.PRICE,
      0
    )
  ) AS Sat
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
  Format(InvoLine.SetupDate, "ww");
