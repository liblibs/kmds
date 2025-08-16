SELECT
  Inv20091031.ItemID,
  Inv20091031.Onhand,
  Item.ID
FROM
  Inv20091031
  LEFT JOIN Item ON Inv20091031.ItemID = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
