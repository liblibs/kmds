DELETE OrdHead.ID
FROM
  OrdHead
WHERE
  (
    (
      (OrdHead.ID)= [parOrdHeadID]
    )
  );
