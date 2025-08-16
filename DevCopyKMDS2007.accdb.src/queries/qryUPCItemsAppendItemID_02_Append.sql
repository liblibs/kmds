INSERT INTO UPCItem (ItemID, ID)
SELECT
  ItemUPCID.ItemID,
  ItemUPCID.UPCID
FROM
  ItemUPCID
  LEFT JOIN UPCItem ON ItemUPCID.UPCID = UPCItem.ID
WHERE
  (
    (
      (UPCItem.ID) Is Null
    )
  );
