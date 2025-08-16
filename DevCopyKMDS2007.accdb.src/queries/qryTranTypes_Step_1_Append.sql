INSERT INTO TranTypes (Code, Description)
SELECT
  TranTypesLocal.Code,
  TranTypesLocal.Description
FROM
  TranTypes
  RIGHT JOIN TranTypesLocal ON TranTypes.Code = TranTypesLocal.Code
WHERE
  (
    (
      (TranTypes.Code) Is Null
    )
  );
