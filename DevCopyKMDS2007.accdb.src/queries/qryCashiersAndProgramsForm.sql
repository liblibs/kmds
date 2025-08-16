SELECT
  [FirstLetter] & [Switchboard Items].ItemText AS MenuOption,
  [Switchboard Items].ItemText,
  [Switchboard Items].Command,
  [Switchboard Items].Argument
FROM
  [Switchboard Items]
  INNER JOIN qrySwitchBoardFirstLetter ON [Switchboard Items].SwitchboardID = qrySwitchBoardFirstLetter.SwitchboardID
WHERE
  (
    (
      ([Switchboard Items].Command)= 3
    )
  )
ORDER BY
  [FirstLetter] & [Switchboard Items].ItemText;
