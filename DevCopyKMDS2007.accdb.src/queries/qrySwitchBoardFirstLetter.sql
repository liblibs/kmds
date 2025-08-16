SELECT
  [Switchboard Items].SwitchboardID,
  [Switchboard Items].ItemNumber,
  [Switchboard Items].ItemText,
  [Switchboard Items].Command,
  [Switchboard Items].Argument,
  Left([ItemText], 1) AS FirstLetter
FROM
  [Switchboard Items]
WHERE
  (
    (
      ([Switchboard Items].Command)= 0
    )
  );
