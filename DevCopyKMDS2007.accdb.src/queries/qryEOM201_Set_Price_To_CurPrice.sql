UPDATE
  Item
  INNER JOIN Newprice ON Item.ID = Newprice.Code
SET
  Item.Price = CCur([NewPrice])* 0.01
WHERE
  (
    (
      (Item.Price)<> CCur([NewPrice])* 0.01
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  );
