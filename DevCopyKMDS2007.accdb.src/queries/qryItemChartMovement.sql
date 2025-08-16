SELECT
  ItemMovement.Code,
  ItemMovement.Description,
  Item.Onhand,
  Item.OrderAt,
  Item.OrderTo,
  ItemMovement.CurrMM,
  ItemMovement.Jan,
  ItemMovement.Feb,
  ItemMovement.Mar,
  ItemMovement.Apr,
  ItemMovement.May,
  ItemMovement.Jun,
  ItemMovement.Jul,
  ItemMovement.Aug,
  ItemMovement.Sep,
  ItemMovement.Oct,
  ItemMovement.Nov,
  ItemMovement.Dec
FROM
  ItemMovement
  INNER JOIN Item ON ItemMovement.Code = Item.ID
WHERE
  (
    (
      (ItemMovement.Code)= Forms!ItemWithSalesTotalChart!ID
    )
    And (
      (ItemMovement.Description)= "  SALES -- TOTAL"
    )
  );
