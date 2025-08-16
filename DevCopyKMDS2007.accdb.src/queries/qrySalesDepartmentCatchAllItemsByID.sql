PARAMETERS parSalesDepartmentID Long;
SELECT
  SalesDepartment.ID,
  SalesDepartment.CatchAllItemID
FROM
  SalesDepartment
  INNER JOIN Item ON SalesDepartment.CatchAllItemID = Item.ID
WHERE
  (
    (
      (SalesDepartment.ID)= [parSalesDepartmentID]
    )
  );
