DELETE Newprice.Code,
Newprice.Price
FROM
  Newprice
WHERE
  (
    (
      (Newprice.Price) Is Null
    )
  )
  OR (
    (
      (Newprice.Code) Is Null
    )
  );
