DELETE Len([id]) AS Lid
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([id])
      ) Between 4
      And 7
    )
  );
