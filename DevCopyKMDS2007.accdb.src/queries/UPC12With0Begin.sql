SELECT
  Len([ID]) AS LID,
  Sum(1) AS [COUNT],
  Mid([ID], 2, 10) AS ID10,
  checkdigit11(
    "0" & Mid([ID], 2, 10)
  ) AS cd
FROM
  UPCItem
GROUP BY
  Len([ID]),
  Mid([ID], 2, 10),
  checkdigit11(
    "0" & Mid([ID], 2, 10)
  )
HAVING
  (
    (
      (
        Len([ID])
      )= 12
    )
    AND (
      (
        Sum(1)
      )= 2
    )
  );
