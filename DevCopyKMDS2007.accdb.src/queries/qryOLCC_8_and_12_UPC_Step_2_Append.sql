INSERT INTO OLCC_8_and_12_UPC (
  Code, lid, ThirdDigit, UPCCode, upc12
)
SELECT
  dbo_OLCCScanCode.Code,
  Len([dbo_OLCCScanCode].[UPCCode]) AS lid,
  qryOLCC_8_and_12_UPC_FirstChar.ThirdDigit,
  dbo_OLCCScanCode.UPCCode,
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
