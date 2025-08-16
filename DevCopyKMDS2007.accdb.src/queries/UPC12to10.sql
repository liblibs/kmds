SELECT
  Len([ID]) AS LID,
  1 AS [COUNT],
  UPCItem.ID,
  Mid([ID], 2, 10) AS ID10
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([ID])
      )= 12
    )
  );
