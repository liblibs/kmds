INSERT INTO UPCItem (ItemID, ID)
SELECT
  qryMetrologicForOLCCShelfTags.ID,
  qryMetrologicForOLCCShelfTags.NewUPC
FROM
  qryMetrologicForOLCCShelfTags
  LEFT JOIN UPCItem ON qryMetrologicForOLCCShelfTags.NewUPC = UPCItem.ID
WHERE
  (
    (
      (UPCItem.ItemID) Is Null
    )
  );
