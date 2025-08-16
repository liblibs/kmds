UPDATE
  dbo_OLCCScanCode
SET
  dbo_OLCCScanCode.UPCCodeKMDS = [upccode]
WHERE
  (
    (
      (
        Len([upccode])
      )= 10
    )
  );
