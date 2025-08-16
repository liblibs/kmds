SELECT
  ItemMovement.Code,
  ItemMovement.Description,
  IIf(
    IsNull([OnOrder]),
    0,
    [OnOrder]
  ) AS Ordered,
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
  (
    ItemMovement
    INNER JOIN qryItemOnOrder ON ItemMovement.Code = qryItemOnOrder.ItemID
  )
  INNER JOIN Item ON ItemMovement.Code = Item.ID
WHERE
  (
    (
      (ItemMovement.Code)= Forms!PurchaseOrdersOnFile!PurchaseOrdersOnFileSub.Form!ItemID
    )
    And (
      (ItemMovement.Description)= "  SALES -- TOTAL"
    )
  );
