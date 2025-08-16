DELETE Case_quantities.*,
Case_quantities.CreateDate
FROM
  Case_quantities
WHERE
  (
    (
      (Case_quantities.CreateDate) Is Null
    )
  );
