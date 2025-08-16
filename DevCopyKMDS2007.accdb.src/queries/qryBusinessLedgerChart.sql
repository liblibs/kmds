SELECT
  Format(
    involine.SetupDate, "mm/dd/yyyy"
  ) AS BusDate,
  Sum(InvoLine.Price) AS LiquorSales
FROM
  (
    (
      (
        InvoLine
        INNER JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
      )
      INNER JOIN Item ON InvoLine.Code = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN SalesGroup ON SalesDepartment.SalesGroupID = SalesGroup.ID
WHERE
  (
    (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (SalesGroup.Description)= "Liquor"
    )
    And (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
  )
GROUP BY
  Format(
    involine.SetupDate, "mm/dd/yyyy"
  );
