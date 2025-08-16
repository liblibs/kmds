UPDATE
  TranTypes
  INNER JOIN TranTypesLocal ON TranTypes.Code = TranTypesLocal.Code
SET
  TranTypes.Description = TranTypesLocal.description;
