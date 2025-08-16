SELECT
  Item.Location,
  Item.ID,
  Item.Description,
  Size.Description AS SizeD,
  [onhand] + [TransferredToday] - [SoldTodayCr1] - [SoldTodayCr2] - [SoldTodayCr3] - [SoldTodayCr4] AS OnTheShelf,
  Item.VendorID,
  qryItemMovementWithTransfersReversed.CurrMMQty,
  qryItemMovementWithTransfersReversed.JanQty,
  qryItemMovementWithTransfersReversed.FebQty,
  qryItemMovementWithTransfersReversed.MarQty,
  qryItemMovementWithTransfersReversed.AprQty,
  qryItemMovementWithTransfersReversed.MayQty,
  qryItemMovementWithTransfersReversed.JunQty,
  qryItemMovementWithTransfersReversed.JulQty,
  qryItemMovementWithTransfersReversed.AugQty,
  qryItemMovementWithTransfersReversed.SepQty,
  qryItemMovementWithTransfersReversed.OctQty,
  qryItemMovementWithTransfersReversed.NovQty,
  qryItemMovementWithTransfersReversed.DecQty,
  qryItemMovementWithTransfersReversed.DecQty,
  qryItemMovementWithTransfersReversed.DecQty,
  qryItemMovementWithTransfersReversed.YtdQ
FROM
  (
    Item
    INNER JOIN [Size] ON Item.SizeID = Size.ID
  )
  INNER JOIN qryItemMovementWithTransfersReversed ON Item.ID = qryItemMovementWithTransfersReversed.Code
WHERE
  (
    (
      (Item.Location) Between Forms!PrintInventoryReport!cboLocationBegin
      And Forms!PrintInventoryReport!cboLocationEnd
    )
    And (
      (Item.VendorID) Between Forms!PrintInventoryReport!cboVendorIDBegin
      And Forms!PrintInventoryReport!cboVendorIDEnd
    )
  )
ORDER BY
  Item.Location,
  Item.Description;
