SELECT
  dbo_Vendor.ID,
  dbo_Vendor.Name,
  qryItemByVendorCount.Count,
  dbo_Vendor.POMethodDefault AS IsNotObsolete
FROM
  dbo_Vendor
  INNER JOIN qryItemByVendorCount ON dbo_Vendor.ID = qryItemByVendorCount.VendorID
WHERE
  (
    (
      (dbo_Vendor.POMethodDefault)<>-1
    )
  )
ORDER BY
  dbo_Vendor.Name;
