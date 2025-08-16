INSERT INTO dbo_OLCCScanCode (Code, UPCCode)
SELECT
  dbo_OLCCScanCode.Code,
  checkdigit11(
    [thirddigit] & [dbo_OLCCScanCode].[upccode]
  ) AS upc12
FROM
  dbo_OLCCScanCode
  INNER JOIN qryOLCC_8_and_12_UPC_FirstChar ON dbo_OLCCScanCode.Code = qryOLCC_8_and_12_UPC_FirstChar.Code
WHERE
  (
    (
      (
        Len([dbo_OLCCScanCode].[UPCCode])
      )= 10
    )
  );
