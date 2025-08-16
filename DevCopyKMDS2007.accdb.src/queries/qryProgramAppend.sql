INSERT INTO Programs (Name)
SELECT
  [Switchboard Items].Argument
FROM
  [Switchboard Items]
  LEFT JOIN Programs ON [Switchboard Items].Argument = Programs.Name
WHERE
  (
    (
      ([Switchboard Items].Command)= 3
    )
    AND (
      (Programs.Name) Is Null
    )
  );
