SELECT
  OLCC_8_and_12_UPC.Code,
  UPCItem.ItemID,
  OLCC_8_and_12_UPC.upc12,
  OLCC_8_and_12_UPC.UPCCode
FROM
  OLCC_8_and_12_UPC
  INNER JOIN UPCItem ON OLCC_8_and_12_UPC.upc12 = UPCItem.ID
WHERE
  (
    (
      (OLCC_8_and_12_UPC.Code)<> [UPCItem].[itemid]
    )
  );
