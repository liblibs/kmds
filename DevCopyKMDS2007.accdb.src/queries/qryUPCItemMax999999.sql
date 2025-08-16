SELECT
  Max(UPCItem.ID) AS MaxOfID
FROM
  UPCItem
HAVING
  (
    (
      (
        Max(UPCItem.ID)
      )> "999999"
    )
    AND (
      (
        Len([ID])
      )= 12
    )
  );
