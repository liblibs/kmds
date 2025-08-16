INSERT INTO upcitem (ItemID, ID)
SELECT
  qryUPCVendor5.ItemID,
  checkdigit11(
    [Vendor6Digit] & Mid([id], 7, 5)
  ) AS NewUPC
FROM
  qryUPCVendor5
  INNER JOIN UPCConvert5to6 ON qryUPCVendor5.Vendor5U = UPCConvert5to6.Vendor5Digit
ORDER BY
  checkdigit11(
    [Vendor6Digit] & Mid([id], 7, 5)
  );
