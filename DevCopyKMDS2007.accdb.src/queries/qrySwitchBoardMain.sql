SELECT
  [Switchboard Items].SwitchboardID,
  [Switchboard Items].ItemNumber,
  [Switchboard Items].ItemText,
  [Switchboard Items].Command,
  [Switchboard Items].Argument
FROM
  [Switchboard Items]
WHERE
  (
    (
      (
        [Switchboard Items].SwitchboardID
      )= 1
    )
    AND (
      ([Switchboard Items].Command)= 1
    )
  );
