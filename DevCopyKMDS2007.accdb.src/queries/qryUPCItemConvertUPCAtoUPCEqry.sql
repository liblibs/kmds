SELECT
  UPCItem.ID,
  UPCItem.ItemID,
  Len([ID]) AS LID,
  IsNumeric([ID]) AS LUN
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([ID])
      )= 10
    )
    AND (
      (
        IsNumeric([ID])
      )= True
    )
  );
