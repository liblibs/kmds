SELECT
  UPCItem.ID,
  UPCItem.ItemID AS KMDSItem,
  Item_1.Description AS KMDSDesc,
  Item_1.Onhand,
  dbo_OLCCScanCode.Code AS OLCCItem,
  Item.Description AS OLCCDesc,
  dbo_OLCCScanCode.SetupDate
FROM
  (
    (
      dbo_OLCCScanCode
      INNER JOIN UPCItem ON dbo_OLCCScanCode.UPCCodeKMDS = UPCItem.ID
    )
    INNER JOIN Item AS Item_1 ON UPCItem.ItemID = Item_1.ID
  )
  INNER JOIN Item ON dbo_OLCCScanCode.Code = Item.ID
WHERE
  (
    (
      (dbo_OLCCScanCode.Code)<> [upcitem].[ItemID]
    )
  )
ORDER BY
  Item_1.Description;
