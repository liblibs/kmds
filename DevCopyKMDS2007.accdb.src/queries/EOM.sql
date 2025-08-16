SELECT
  MSysObjects.Name,
  MSysObjects.Type
FROM
  MSysObjects
WHERE
  (
    (
      (MSysObjects.Name) Like 'qryeom*'
    )
  )
ORDER BY
  MSysObjects.Name;
