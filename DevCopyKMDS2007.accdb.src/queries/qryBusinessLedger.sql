SELECT
  Format(
    involine.SetupDate, "mm/dd/yyyy"
  ) AS BusDate,
  Sum(
    IIf(
      [Transtype] = "S", Involine.Price,
      0
    )
  ) AS SalesCounter,
  Sum(
    IIf(
      [Transtype] = "C", Involine.Price,
      0
    )
  ) AS SalesClubs,
  Sum(
    IIf(
      [Transtype] = "P", Involine.Price,
      0
    )
  ) AS Purchases,
  Sum(
    IIf(
      [Transtype] = "X"
      And Involine.Price > 0,
      Involine.Price, 0
    )
  ) AS TransfersIn,
  Sum(
    IIf(
      [Transtype] = "X"
      And Involine.Price < 0,
      Involine.Price, 0
    )
  ) AS TransfersOut,
  Sum(
    IIf(
      [Transtype] = "A", Involine.Price,
      0
    )
  ) AS Adjust
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
  )
GROUP BY
  Format(
    involine.SetupDate, "mm/dd/yyyy"
  );
