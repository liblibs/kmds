SELECT
  InvoLine.SalesGroup,
  Vendor.Name AS VendorName,
  SalesDepartment.ID AS Department,
  IIf(
    involine.Transtype = 'C', "Club Sales",
    SalesDepartment.Description
  ) AS DepartmentDesc,
  Item.PriceGroupID,
  PriceGroup.Description AS PriceGroupDescription,
  Item.Description,
  Item.ID,
  InvoLine.SetupDate,
  InvoLine.Quantity AS TotalQ,
  InvoLine.Price AS TotalP,
  InvoLine.TransType,
  InvoLine.InputCode,
  InvoLine.InvoHeadID AS PONumber,
  InvoLine.Cost
FROM
  (
    (
      InvoLine
      INNER JOIN (
        SalesGroup
        INNER JOIN (
          SalesDepartment
          INNER JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
        ) ON SalesGroup.ID = SalesDepartment.SalesGroupID
      ) ON InvoLine.Code = Item.ID
    )
    INNER JOIN Vendor ON Item.VendorID = Vendor.ID
  )
  INNER JOIN PriceGroup ON Item.PriceGroupID = PriceGroup.ID
WHERE
  (
    (
      (InvoLine.SalesGroup) Between Forms!PrintInventoryReport!cboVendorIDBegin
      And Forms!PrintInventoryReport!cboVendorIDEnd
    )
    And (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (PriceGroup.Description) Between Forms!PrintInventoryReport!cboPriceGroupIDBegin
      And Forms!PrintInventoryReport!cboPriceGroupIDEnd
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (InvoLine.TransType)= "P"
    )
  )
ORDER BY
  SalesDepartment.ID,
  Item.PriceGroupID,
  Item.Description,
  InvoLine.SetupDate;
