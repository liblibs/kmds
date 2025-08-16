SELECT
  Item.SalesDepartmentID AS SalesDeptID,
  SalesDepartment.Description AS SalesDept,
  Item.VendorID,
  Item.Description,
  Item.ID,
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
      (Item.Price)> 0
    )
    AND (
      (Item.PriceGroupID)= 0
    )
  );
