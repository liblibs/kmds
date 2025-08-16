SELECT
  Item.VendorID,
  Item.Description,
  Item.ID,
  Item.SalesDepartmentID AS SalesDeptID,
  SalesDepartment.Description AS SalesDept,
  SubCategory.Description AS Category,
  Item.Price,
  Item.Cost
FROM
  (
    (
      Item
      INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN Vendor ON Item.VendorID = Vendor.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)> 2
    )
    AND (
      (Item.PriceGroupID)= 0
    )
  );
