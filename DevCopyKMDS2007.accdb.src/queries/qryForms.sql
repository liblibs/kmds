SELECT
  MSysObjects.Name,
  MSysObjects.Type
FROM
  MSysObjects
WHERE
  (
    (
      (MSysObjects.Name)<> "qryForms"
    )
    AND (
      (MSysObjects.Type)=-32768
    )
  )
ORDER BY
  MSysObjects.Name;
