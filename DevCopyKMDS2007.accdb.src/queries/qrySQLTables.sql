SELECT
  MSysObjects.Connect,
  MSysObjects.Name AS TableName
FROM
  MSysObjects
WHERE
  (
    (
      (MSysObjects.Connect) Is Not Null
    )
  );
