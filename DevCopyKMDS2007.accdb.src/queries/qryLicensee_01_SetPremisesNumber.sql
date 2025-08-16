UPDATE
  LocLicensee
SET
  LocLicensee.PremisesNumber = CLng(
    [loclicensee].[PremisesNumberImported]
  )
WHERE
  (
    (
      (
        LocLicensee.PremisesNumberImported
      ) Like "000*"
    )
  );
