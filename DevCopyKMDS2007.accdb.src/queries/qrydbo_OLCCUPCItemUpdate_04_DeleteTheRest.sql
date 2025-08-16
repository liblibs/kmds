DELETE dbo_OLCCScanCode.UPCCodeKMDS,
dbo_OLCCScanCode.*
FROM
  dbo_OLCCScanCode
WHERE
  (
    (
      (dbo_OLCCScanCode.UPCCodeKMDS) Is Null
    )
  );
