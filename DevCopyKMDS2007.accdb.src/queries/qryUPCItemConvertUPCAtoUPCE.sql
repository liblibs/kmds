INSERT INTO UPCItem (ItemID, ID)
SELECT
  qryUPCItemConvertUPCAtoUPCEqry.ItemID,
  upcatoe([id]) AS NewID
FROM
  qryUPCItemConvertUPCAtoUPCEqry
WHERE
  (
    (
      (
        upcatoe([id])
      )<> [id]
    )
  );
