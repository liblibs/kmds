INSERT INTO UPCItem (ID, ItemID)
SELECT
  dbo_OLCCScanCode.UPCCode,
  dbo_OLCCScanCode.Code
FROM
  (
    dbo_OLCCScanCode
    INNER JOIN Item ON dbo_OLCCScanCode.Code = Item.ID
  )
  LEFT JOIN UPCItem ON dbo_OLCCScanCode.UPCCode = UPCItem.ID
WHERE
  (
    (
      (UPCItem.ID) Is Null
    )
  );
