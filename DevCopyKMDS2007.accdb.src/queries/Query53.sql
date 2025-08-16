SELECT
  SalesDepartment.ID,
  SalesDepartment.Description,
  Item.ID,
  qryUPCMRU.ID AS UPC,
  Item.Description,
  Item.Price,
  Item.Cost,
  Item.Deposit
FROM
  (
    Item
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN qryUPCMRU ON Item.ID = qryUPCMRU.ItemID
WHERE
  (
    (
      (SalesDepartment.ID)> 2
    )
  )
ORDER BY
  SalesDepartment.ID,
  Item.Description;
