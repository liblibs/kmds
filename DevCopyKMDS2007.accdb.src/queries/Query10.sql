INSERT INTO item (
  ID, Description, SalesDepartmentID,
  Onhand, Cost, Price
)
SELECT
  MeadowlandInventoryValuex.NewItem,
  MeadowlandInventoryValuex.description,
  MeadowlandInventoryValuex.SalesDpt AS SalesDepartmentID,
  MeadowlandInventoryValuex.Onhand,
  MeadowlandInventoryValuex.Cost,
  MeadowlandInventoryValuex.Price
FROM
  MeadowlandInventoryValuex
  LEFT JOIN SalesDepartment ON MeadowlandInventoryValuex.SalesDpt = SalesDepartment.ID
WHERE
  (
    (
      (
        MeadowlandInventoryValuex.NewItem
      )> 0
    )
  );
