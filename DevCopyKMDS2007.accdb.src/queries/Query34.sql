UPDATE
  crescentUPC
  INNER JOIN UPCItem ON crescentUPC.ID = UPCItem.ID
SET
  UPCItem.ID = UPCECheckdigit(
    Trim([crescentUPC].[ID])
  )
WHERE
  (
    (
      (crescentUPC.lid)<> [crescentUPC].[Itemid]
      And (crescentUPC.lid)> 0
    )
  );
