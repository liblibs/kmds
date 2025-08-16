INSERT INTO ItemMovement (Code, Description, CurrMM)
SELECT
  ItemMovementCurrentMonth.Code,
  ItemMovementCurrentMonth.Description,
  ItemMovementCurrentMonth.CurMQty AS CurrMM
FROM
  ItemMovementCurrentMonth
  LEFT JOIN ItemMovement ON (
    ItemMovementCurrentMonth.Description = ItemMovement.Description
  )
  AND (
    ItemMovementCurrentMonth.Code = ItemMovement.Code
  )
WHERE
  (
    (
      (ItemMovement.Description) Is Null
    )
  );
