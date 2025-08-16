INSERT INTO Licensee (
  PremisesNumber, TradeName, Street,
  City, Zip
)
SELECT
  LocLicensee.PremisesNumber,
  LocLicensee.TradeName,
  LocLicensee.Street,
  LocLicensee.City,
  LocLicensee.Zip
FROM
  LocLicensee
  LEFT JOIN Licensee ON LocLicensee.PremisesNumber = Licensee.PremisesNumber
WHERE
  (
    (
      (LocLicensee.PremisesNumber) Is Not Null
    )
    AND (
      (Licensee.PremisesNumber) Is Null
    )
  );
