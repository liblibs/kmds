DELETE Item.AvgTurns
FROM
  Item
WHERE
  (
    (
      (Item.AvgTurns)= 999
    )
  );
