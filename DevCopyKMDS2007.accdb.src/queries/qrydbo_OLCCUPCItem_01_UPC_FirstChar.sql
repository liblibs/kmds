SELECT
  dbo_OLCCScanCode.ID,
  dbo_OLCCScanCode.Code,
  dbo_OLCCScanCode.UPCCode,
  Len([UPCCode]) AS LID,
  Mid([upccode], 3, 1) AS ThirdDigit
FROM
  dbo_OLCCScanCode
WHERE
  (
    (
      (dbo_OLCCScanCode.UPCCode) Like "10*"
    )
    AND (
      (
        Len([UPCCode])
      )= 14
    )
  );
