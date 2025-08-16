PARAMETERS parInvoiceNo Long;
SELECT
  OrdHead.ID,
  OrdHead.ClubID,
  OrdHead.Order,
  OrdHead.InvoiceNo,
  OrdHead.OrdDate,
  OrdHead.InvDate,
  Club.Name AS ClubName,
  Club.Contact AS ClubContact,
  Club.Address1 AS ClubAddress1,
  Club.Address2 AS ClubAddress2,
  [City] & ", " & [ST] & "  " & [Zip] AS ClubCityST,
  Club.Phone1 AS ClubPhone1,
  OrderLine.LineID,
  OrderLine.ItemID,
  Item.Description,
  Size.Description AS SizeDesc,
  OrderLine.QuantityShipped,
  OrderLine.Price,
  OrderLine.Cost,
  OrdHead.OrderNotes,
  OrderLine.SalePrice,
  Size.Oz,
  CCur(OrderLine.price / [oz]) AS PricePerOZ
FROM
  (
    Club
    INNER JOIN OrdHead ON Club.ID = OrdHead.ClubID
  )
  INNER JOIN (
    (
      Item
      LEFT JOIN [Size] ON Item.SizeID = Size.ID
    )
    INNER JOIN OrderLine ON Item.ID = OrderLine.ItemID
  ) ON OrdHead.ID = OrderLine.OrdheadID
WHERE
  (
    (
      (OrdHead.InvoiceNo)= [parInvoiceNo]
    )
    AND (
      (OrderLine.QuantityShipped)<> 0
    )
  );
