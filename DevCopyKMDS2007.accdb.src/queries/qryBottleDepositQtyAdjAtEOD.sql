SELECT
  Settings.KeyName,
  UCase([Value]) AS UV
FROM
  Settings
WHERE
  (
    (
      (Settings.KeyName)= "BottleDepositQtyAdjAtEOD"
    )
    AND (
      (
        UCase([Value])
      )= "TRUE"
    )
  );
