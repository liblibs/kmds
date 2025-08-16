UPDATE
  (
    Item
    LEFT JOIN InvoLine ON Item.ID = InvoLine.Code
  )
  LEFT JOIN ClubOrderGuides ON Item.ID = ClubOrderGuides.ItemID
SET
  Item.AvgTurns = 999
WHERE
  (
    (
      (InvoLine.Code) Is Null
    )
    AND (
      (Item.Onhand)= 0
    )
    AND (
      (Item.SetupDate)< Now()-60
    )
    AND (
      (ClubOrderGuides.ClubID) Is Null
    )
  );
