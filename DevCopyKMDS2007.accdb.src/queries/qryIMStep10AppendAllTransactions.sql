INSERT INTO ItemMovement (Code, Description)
SELECT
  qryALLItemsAndALLTransTypes.ID,
  qryALLItemsAndALLTransTypes.Description
FROM
  qryALLItemsAndALLTransTypes
  LEFT JOIN ItemMovement ON (
    qryALLItemsAndALLTransTypes.Description = ItemMovement.Description
  )
  AND (
    qryALLItemsAndALLTransTypes.ID = ItemMovement.Code
  )
WHERE
  (
    (
      (ItemMovement.Code) Is Null
    )
  );
