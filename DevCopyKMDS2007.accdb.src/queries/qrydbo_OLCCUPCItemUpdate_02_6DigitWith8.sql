UPDATE
  dbo_OLCCScanCode
SET
  dbo_OLCCScanCode.UPCCodeKMDS = UPCECheckdigit([upccode])
WHERE
  (
    (
      (
        Len([upccode])
      )= 6
    )
  );
