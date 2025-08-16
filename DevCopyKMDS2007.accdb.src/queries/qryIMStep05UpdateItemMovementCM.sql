UPDATE
  ItemMovement
  INNER JOIN ItemMovementCurrentMonth ON (
    ItemMovement.Code = ItemMovementCurrentMonth.Code
  )
  AND (
    ItemMovement.Description = ItemMovementCurrentMonth.Description
  )
SET
  ItemMovement.CurrMM = [CurMQty];
