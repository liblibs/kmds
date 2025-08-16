PARAMETERS parVendorID Long;
SELECT
  Vendor.ID,
  Vendor.OrderCount
FROM
  Vendor
WHERE
  (
    (
      (Vendor.ID)= [parVendorID]
    )
  );
