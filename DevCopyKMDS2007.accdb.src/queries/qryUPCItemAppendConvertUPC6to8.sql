UPDATE
  UPCItem
SET
  UPCItem.ID = upcecheckdigit([id])
WHERE
  (
    (
      (
        CDbl([id])
      )<> [ItemID]
    )
    AND (
      (
        Len([ID])
      )= 6
    )
  );
