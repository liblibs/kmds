UPDATE
  UPCItem
  INNER JOIN upcOld ON UPCItem.ID = upcOld.ID
SET
  UPCItem.SetupDate = #7/27/1956#;
