INSERT INTO dbo_OLCCScanCode (Code, UPCCode)
SELECT
  dbo_OLCCScanCode.Code,
  upcecheckdigit("0" & [upccode]) AS Expr1
FROM
  dbo_OLCCScanCode
WHERE
  (
    (
      (
        (
          Len([upccode])
        )
      )= 6
    )
  );
