SELECT
  SubCategory.Description,
  qryUPCItemCountPerItemID.Count,
  Item.ID,
  Item.Description,
  Item.SalesDepartmentID,
  SalesDepartment.Description AS SDesc
FROM
  (
    (
      Item
      INNER JOIN qryUPCItemCountPerItemID ON Item.ID = qryUPCItemCountPerItemID.ItemID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
ORDER BY
  SubCategory.Description,
  qryUPCItemCountPerItemID.Count DESC,
  Item.ID;
