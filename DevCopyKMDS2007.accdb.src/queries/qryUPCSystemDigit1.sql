UPDATE
  UPCItem
SET
  UPCItem.ID = checkdigit11(
    [parNewSystem] & Mid([ID], 2, 10)
  )
WHERE
  (
    (
      (UPCItem.ID) Like [parOldVendorID]
    )
    AND (
      (
        Len([id])
      )= 12
    )
  );
