INSERT INTO UPCConvert5to6 (Vendor6Digit, Vendor5Digit)
SELECT
  Left([id], 6) AS Vendor6,
  Mid([id], 2, 5) AS Vendor5
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([ID])
      )= 12
    )
    AND (
      (UPCItem.ID)> "2"
    )
  )
GROUP BY
  Left([id], 6),
  Mid([id], 2, 5);
