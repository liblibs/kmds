UPDATE
  [Switchboard Items]
SET
  [Switchboard Items].Command = [Command] + 1
WHERE
  (
    (
      ([Switchboard Items].Argument)= "default"
    )
  );
