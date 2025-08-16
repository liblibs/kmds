SELECT
  Item.ID,
  TranTypes.Description,
  TranTypes.Code
FROM
  Item,
  TranTypes
WHERE
  (
    (
      (TranTypes.Code)<> "T"
      And (TranTypes.Code)<> "R"
      And (TranTypes.Code)<> "D"
    )
  );
