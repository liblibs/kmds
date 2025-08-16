SELECT
  OLCCScanCode.UPCCode,
  Len([UPCCode]) AS lid,
  Mid([upccode], 4, 10) AS UPC10,
  Mid([upccode], 3, 1) AS FirstChar,
  checkdigit11([FirstChar] & [UPC10]) AS upc12,
  OLCCScanCode.Code
FROM
  OLCCScanCode
WHERE
  (
    (
      (
        Len([UPCCode])
      )= 14
    )
  );
