INSERT INTO dbo_OLCCScanCode (Code, UPCCode)
SELECT
  First(OLCCScanCode.Code) AS FirstOfCode,
  OLCCScanCode.UPCCode
FROM
  OLCCScanCode
WHERE
  (
    (
      (
        Len([UPCCode])
      )= 10
      Or (
        Len([UPCCode])
      )= 6
      Or (
        Len([UPCCode])
      )= 14
    )
  )
GROUP BY
  OLCCScanCode.UPCCode;
