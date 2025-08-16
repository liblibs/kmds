UPDATE
  dbo_OLCCScanCode
SET
  dbo_OLCCScanCode.UPCCodeKMDS = checkdigit11(
    Mid([upccode], 3, 11)
  )
WHERE
  (
    (
      (
        Len([upccode])
      )= 14
    )
  );
