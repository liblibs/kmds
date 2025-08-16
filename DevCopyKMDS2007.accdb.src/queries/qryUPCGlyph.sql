SELECT
  qryUPCItemMax.ItemID,
  UPCGlyph([UPCCodeMax]) AS UPCG,
  qryUPCItemMax.UPCCodeMax
FROM
  qryUPCItemMax;
