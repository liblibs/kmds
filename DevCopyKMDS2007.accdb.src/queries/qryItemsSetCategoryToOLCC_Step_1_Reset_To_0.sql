UPDATE
  Item
SET
  Item.SubCategoryID = 0
WHERE
  (
    (
      (Item.SubCategoryID)<> 0
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  );
