INSERT INTO dbo_OLCCScanCode (Code, UPCCode)
SELECT
  dbo_OLCCScanCode.Code,
  checkdigit11(
    [thirddigit] & [dbo_OLCCScanCode].[upccode]
  ) AS upc12
FROM
  qrydbo_OLCCUPCItem_01_UPC_FirstChar
  INNER JOIN dbo_OLCCScanCode ON qrydbo_OLCCUPCItem_01_UPC_FirstChar.Code = dbo_OLCCScanCode.Code
WHERE
  (
    (
      (
        Len([dbo_OLCCScanCode].[UPCCode])
      )= 10
    )
  )
GROUP BY
  dbo_OLCCScanCode.Code,
  checkdigit11(
    [thirddigit] & [dbo_OLCCScanCode].[upccode]
  );
