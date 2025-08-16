SELECT
  CLng([Value]) AS Division_ID
FROM
  Settings
WHERE
  (
    (
      (Settings.KeyName)= "Fintech.Division-IDRelated"
    )
  );
