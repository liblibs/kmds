SELECT
  MetrologicScanbatch.ID,
  UPCItem.ItemID,
  Item.Description,
  MetrologicScanbatch.ScanCode,
  MetrologicScanbatch.Quantity,
  Item.Price,
  MetrologicScanbatch.ScanDate,
  MetrologicScanbatch.ScanTime,
  SalesDepartment.Description AS SalesDptDescr,
  SubCategory.Description AS CategoryDescr
FROM
  (
    (
      (
        MetrologicScanbatch
        INNER JOIN UPCItem ON MetrologicScanbatch.ScanCode = UPCItem.ID
      )
      INNER JOIN Item ON UPCItem.ItemID = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
ORDER BY
  MetrologicScanbatch.ScanDate,
  MetrologicScanbatch.ScanTime;
