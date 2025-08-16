UPDATE
  Licensee
  INNER JOIN LocLicensee ON Licensee.PremisesNumber = LocLicensee.PremisesNumber
SET
  Licensee.TradeName = LocLicensee.TradeName,
  Licensee.Street = LocLicensee.Street,
  Licensee.City = LocLicensee.City,
  Licensee.Zip = LocLicensee.Zip
WHERE
  (
    (
      (Licensee.TradeName)<> LocLicensee.TradeName
    )
  )
  Or (
    (
      (Licensee.Street)<> LocLicensee.Street
    )
  )
  Or (
    (
      (Licensee.City)<> LocLicensee.City
    )
  )
  Or (
    (
      (Licensee.Zip)<> LocLicensee.Zip
    )
  );
