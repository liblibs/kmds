SELECT
  Len([ID]) AS LID,
  1 AS [COUNT],
  UPCItem.ID,
  Mid([ID], 2, 6) AS ID10
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([ID])
      )= 8
    )
  );
