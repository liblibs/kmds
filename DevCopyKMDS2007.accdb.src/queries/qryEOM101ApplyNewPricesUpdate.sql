UPDATE
  Newprice
  INNER JOIN Item ON Newprice.Code = Item.ID
SET
  Item.Price = CCur(Newprice.NewPrice * 0.01),
  Item.ChangeDate = Now()
WHERE
  (
    (
      (Newprice.StartDate)> " "
    )
    AND (
      (Newprice.EndDate)= "000000"
    )
  );
