SELECT
  Prod1.Code,
  Item.ID,
  Prod1.Field1
FROM
  Prod1
  LEFT JOIN Item ON Prod1.Code = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
