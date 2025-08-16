DELETE OrdHead.*,
OrdHead.PaidFlag
FROM
  OrdHead
WHERE
  (
    (
      (OrdHead.PaidFlag)<> 0
    )
  );
