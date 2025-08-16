SELECT
  OLCCScanCode.Code,
  OLCCScanCode.UPCCode AS UPC14,
  Mid([upccode], 4, 10) AS UPC10,
  Mid([upccode], 3, 12) AS UPC12
FROM
  OLCCScanCode
WHERE
  (
    (
      (
        Len([upccode])
      )= 14
    )
  )
ORDER BY
  Len([upccode]);
