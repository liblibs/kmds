UPDATE
  Item
  LEFT JOIN Vendor ON Item.VendorID = Vendor.ID
SET
  Item.VendorID = 0
WHERE
  (
    (
      (Item.VendorID)<> 0
    )
    AND (
      (Vendor.ID) Is Null
    )
  );
