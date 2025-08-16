SELECT
  MSysObjects.Name,
  MSysObjects.Type
FROM
  MSysObjects
WHERE
  (
    (
      (MSysObjects.Name) Like 'qryEOD*'
    )
  )
ORDER BY
  MSysObjects.Name;
