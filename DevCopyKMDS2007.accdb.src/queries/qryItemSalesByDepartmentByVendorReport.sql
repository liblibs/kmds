SELECT
  SalesDepartment.ID AS SalesDptID,
  SalesDepartment.Description AS SalesDptDesc,
  Vendor.ID AS VendorID,
  Vendor.Name AS VendorName,
  IIf(
    IsNull(SubCategory.ID),
    "n/a",
    Subcategory.Description
  ) AS Category,
  Item.ID,
  Item.Description,
  Size.Description AS SizeDesc,
  Item.Onhand,
  Item.Price,
  qryItemsByDaysSupplySoldTotals.SumOfQuantity,
  qryItemsByDaysSupplySoldTotals.SumOfPrice,
  Item.SetupDate,
  Item.DateLastSold,
  qryItemsByDaysSupplySoldTotals.SumOfCost,
  IIf(
    [SumOfPrice] = 0,
    0,
    ([SumOfPrice] - [SumOfCost])/ [SumOfPrice]
  ) AS Profit
FROM
  (
    (
      (
        SalesDepartment
        INNER JOIN (
          Item
          INNER JOIN qryItemsByDaysSupplySoldTotals ON Item.ID = qryItemsByDaysSupplySoldTotals.Code
        ) ON SalesDepartment.ID = Item.SalesDepartmentID
      )
      INNER JOIN [Size] ON Item.SizeID = Size.ID
    )
    LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
  )
  INNER JOIN Vendor ON Item.VendorID = Vendor.ID
WHERE
  (
    (
      (SalesDepartment.ID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (Vendor.ID) Between Forms!PrintInventoryReport!cboVendorIDBegin
      And Forms!PrintInventoryReport!cboVendorIDEnd
    )
  )
ORDER BY
  SalesDepartment.ID,
  Vendor.ID;
