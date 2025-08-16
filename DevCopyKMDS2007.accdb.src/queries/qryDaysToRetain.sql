SELECT
  IIf(
    (
      CLng([Value]) Between 500
      And 1000
    ),
    CLng([Value]),
    500
  ) AS Days
FROM
  Settings
WHERE
  (
    (
      (Settings.KeyName)= "DaysToRetain"
    )
  );
