INSERT INTO OLASItemDetailForOrdering (
  ItemCode, NewItemCode, Description,
  Category, [Size], Status, NoWillCalls,
  MinOrderQty, MaxOrderQty, BottlesPerCase,
  PricePerBottle, QuantityAvailable
)
SELECT
  ItemDetailForOrdering.F1,
  ItemDetailForOrdering.F2,
  ItemDetailForOrdering.F3,
  ItemDetailForOrdering.F4,
  ItemDetailForOrdering.F5,
  ItemDetailForOrdering.F6,
  ItemDetailForOrdering.F7,
  ItemDetailForOrdering.F8,
  ItemDetailForOrdering.F9,
  ItemDetailForOrdering.F10,
  ItemDetailForOrdering.F11,
  ItemDetailForOrdering.F12
FROM
  ItemDetailForOrdering
WHERE
  (
    (
      (ItemDetailForOrdering.F2) Like "999*"
    )
  );
