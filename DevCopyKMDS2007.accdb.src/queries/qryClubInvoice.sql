SELECT
  IIf([SalesDepartmentID] = 1,-1, 0) AS LiquorItem,
  OrdHead.ID,
  OrdHead.ClubID,
  OrdHead.Order,
  OrdHead.InvoiceNo,
  OrdHead.OrdDate,
  OrdHead.InvDate,
  qryClubOrderLineTotals.Quantity,
  qryClubOrderLineTotals.ExtCost,
  qryClubOrderLineTotals.LiquorExtPrice,
  qryClubOrderLineTotals.LiquorDiscount,
  [LiquorExtPrice] - [LiquorDiscount] AS LiquorNetPrice,
  qryClubOrderLineTotals.NonLiquorExtPrice,
  qryClubOrderLineTotals.NonLiquorDiscount,
  [NonLiquorExtPrice] - [NonLiquorDiscount] AS NonLiquorNetPrice,
  qryClubOrderLineTotals.NonLiquor,
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
  OrderLine.Buydown,
  OrderLine.Price - [BuyDown] AS NetPrice,
  IIf(
    (
      IsNull(size.Oz)
      Or size.oz = 0
    ),
    1,
    size.oz
  ) AS Ounces,
  IIf(
    item.salesdepartmentId <> 1,
    orderline.price,
    CCur(
      (
        (OrderLine.price / [Ounces])* Club.OuncesColumn
      )* 0.95
    )
  ) AS PricePerOZ,
  Club.Phone2 AS ClubFax,
  Item.SalesDepartmentID,
  SubCategory.Description AS CatDesc,
  Item.Location,
  Club.OuncesColumn,
  IIf(
    item.salesdepartmentId <> 1,
    (orderline.price - [buydown]),
    CCur(OrderLine.price * 0.95)
  ) AS DiscountedPrice
FROM
  (
    (
      Club
      INNER JOIN (
        OrdHead
        INNER JOIN qryClubOrderLineTotals ON OrdHead.ID = qryClubOrderLineTotals.OrdheadID
      ) ON Club.ID = OrdHead.ClubID
    )
    INNER JOIN (
      (
        Item
        LEFT JOIN [Size] ON Item.SizeID = Size.ID
      )
      INNER JOIN OrderLine ON Item.ID = OrderLine.ItemID
    ) ON OrdHead.ID = OrderLine.OrdheadID
  )
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (OrdHead.InvoiceNo)= Forms!ClubOrderPrintDialog!txtInvoiceNo
    )
    And (
      (OrderLine.QuantityShipped)<> 0
    )
  );
