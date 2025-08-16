SELECT
  MSysObjects.Name,
  MSysObjects.Type
FROM
  MSysObjects
WHERE
  (
    (
      (MSysObjects.Name) Like 'qryeod*'
    )
  )
ORDER BY
  MSysObjects.Name;
