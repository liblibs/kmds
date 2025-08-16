SELECT
  dbo_OLCCScanCode.Code,
  dbo_OLCCScanCode.UPCCodeKMDS,
  Len([UPCCodeKMDS]) AS lid
FROM
  dbo_OLCCScanCode
WHERE
  (
    (
      (
        Len([UPCCodeKMDS])
      )= 12
      Or (
        Len([UPCCodeKMDS])
      )= 8
    )
  );
