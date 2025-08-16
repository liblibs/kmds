INSERT INTO upcitem (ItemID, ID)
SELECT
  qryUPCItemConvertUPCAtoUPCEqry.ItemID,
  checkdigit11("0" & [id]) AS newid
FROM
  qryUPCItemConvertUPCAtoUPCEqry;
