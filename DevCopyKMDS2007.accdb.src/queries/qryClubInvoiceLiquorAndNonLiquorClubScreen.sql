SELECT
  qryClubInvoiceForClubScreen.LiquorItem,
  Sysfile.*,
  qryClubInvoiceForClubScreen.ClubID,
  qryClubInvoiceForClubScreen.InvoiceNo AS [Order],
  qryClubInvoiceForClubScreen.InvoiceNo,
  qryClubInvoiceForClubScreen.OrdDate,
  qryClubInvoiceForClubScreen.InvDate,
  qryClubInvoiceForClubScreen.LiquorExtPrice,
  qryClubInvoiceForClubScreen.LiquorDiscount,
  qryClubInvoiceForClubScreen.LiquorNetPrice,
  qryClubInvoiceForClubScreen.NonLiquorExtPrice,
  qryClubInvoiceForClubScreen.NonLiquorDiscount,
  qryClubInvoiceForClubScreen.NonLiquorNetPrice,
  [LiquorExtPrice] + [NonLiquorExtPrice] AS ExtPrice,
  [LiquorDiscount] + [NonLiquorDiscount] AS Discount,
  (
    [LiquorExtPrice] + [NonLiquorExtPrice]
  )-(
    [LiquorDiscount] + [NonLiquorDiscount]
  ) AS ExtNetPrice,
  qryClubInvoiceForClubScreen.NonLiquor,
  qryClubInvoiceForClubScreen.ClubName,
  qryClubInvoiceForClubScreen.ClubContact,
  qryClubInvoiceForClubScreen.ClubAddress1,
  qryClubInvoiceForClubScreen.ClubAddress2,
  qryClubInvoiceForClubScreen.ClubCityST,
  qryClubInvoiceForClubScreen.LineID,
  qryClubInvoiceForClubScreen.ItemID,
  qryClubInvoiceForClubScreen.Description,
  qryClubInvoiceForClubScreen.SizeDesc,
  qryClubInvoiceForClubScreen.QuantityShipped,
  [Netprice] / [QuantityShipped] AS UNetPrice,
  qryClubInvoiceForClubScreen.NetPrice AS XPrice,
  qryClubInvoiceForClubScreen.Cost AS XCost,
  [Cost] / [QuantityShipped] AS UCost,
  qryClubInvoiceForClubScreen.ClubNotes AS OrderNotes,
  qryClubInvoiceForClubScreen.SalePrice,
  qryClubInvoiceForClubScreen.PricePerOZ,
  qryClubInvoiceForClubScreen.ClubFax,
  qryClubInvoiceForClubScreen.SalesDepartmentID,
  qryClubInvoiceForClubScreen.CatDesc,
  qryClubInvoiceForClubScreen.Location,
  qryClubInvoiceForClubScreen.InvoiceNumber
FROM
  Sysfile,
  qryClubInvoiceForClubScreen
ORDER BY
  qryClubInvoiceForClubScreen.LineID;
