SELECT
  Vendor.*
FROM
  Vendor
WHERE
  (
    (
      (Vendor.POMethodDefault)>-1
    )
  )
ORDER BY
  Vendor.Name;
