UPDATE
  dbo_OLCCScanCode
SET
  dbo_OLCCScanCode.UPCCodeKMDS = [UPCCode]
WHERE
  (
    (
      (
        Len([UPCCode])
      )= 8
      Or (
        Len([UPCCode])
      )= 12
    )
  );
