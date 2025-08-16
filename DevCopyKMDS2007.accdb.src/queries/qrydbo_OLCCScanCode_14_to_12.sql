UPDATE
  dbo_OLCCScanCode
SET
  dbo_OLCCScanCode.UPCCodeKMDS = CHECKDIGIT11(
    Mid([UPCCODE], 3, 11)
  )
WHERE
  (
    (
      (
        Len([UPCCODE])
      )= 14
    )
  );
