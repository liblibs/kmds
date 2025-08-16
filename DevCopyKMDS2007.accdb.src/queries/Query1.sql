SELECT
  [Address1] & ", " & [city] & ", " & [st] & " " & [zip] AS ad,
  Club.City,
  Club.St,
  Club.Zip,
  Club.Name
FROM
  Club;
