SELECT
  [Switchboard Items].ItemText,
  [Switchboard Items].SwitchboardID,
  [Switchboard Items].ItemNumber,
  [Switchboard Items].Command,
  CInt([Argument]) AS Arg
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
  )
ORDER BY
  [Switchboard Items].ItemText;
