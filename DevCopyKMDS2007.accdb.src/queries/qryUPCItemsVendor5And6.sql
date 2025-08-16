SELECT
  Len([id]) AS lid,
  Left([id], 6) AS Vendor,
  Mid([id], 2, 5) AS Vendor5
FROM
  UPCItemGH
WHERE
  (
    (
      ([UPCItemGH].[ID])> "19"
    )
  )
GROUP BY
  Len([id]),
  Left([id], 6),
  Mid([id], 2, 5)
HAVING
  (
    (
      (
        Len([id])
      )= 12
    )
  );
