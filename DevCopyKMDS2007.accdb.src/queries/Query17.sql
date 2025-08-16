INSERT INTO UPCItem (ItemID, ID)
SELECT
  MeadowlandBigNumbers.ID,
  MeadowlandInventoryReport13.NewUPC2
FROM
  (
    (
      MeadowlandBigNumbers
      INNER JOIN MeadowlandInventoryValuex ON MeadowlandBigNumbers.ItemNo = MeadowlandInventoryValuex.ItemNo
    )
    INNER JOIN MeadowlandInventoryReport13 ON MeadowlandInventoryValuex.ItemNo = MeadowlandInventoryReport13.Item
  )
  INNER JOIN SalesDepartment ON MeadowlandInventoryValuex.SalesDpt = SalesDepartment.ID
WHERE
  (
    (
      (
        MeadowlandInventoryReport13.NewUPC2
      ) Is Not Null
    )
  );
