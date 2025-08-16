SELECT
  qryClubInvoice.LiquorItem,
  Sysfile.*,
  qryClubInvoice.ClubID,
  qryClubInvoice.Order,
  qryClubInvoice.InvoiceNo,
  qryClubInvoice.OrdDate,
  qryClubInvoice.InvDate,
  qryClubInvoice.LiquorExtPrice,
  qryClubInvoice.LiquorDiscount,
  qryClubInvoice.LiquorNetPrice,
  qryClubInvoice.NonLiquorExtPrice,
  qryClubInvoice.NonLiquorDiscount,
  qryClubInvoice.NonLiquorNetPrice,
  [LiquorExtPrice] + [NonLiquorExtPrice] AS ExtPrice,
  [LiquorDiscount] + [NonLiquorDiscount] AS Discount,
  (
    [LiquorExtPrice] + [NonLiquorExtPrice]
  )-(
    [LiquorDiscount] + [NonLiquorDiscount]
  ) AS ExtNetPrice,
  qryClubInvoice.NonLiquor,
  qryClubInvoice.ClubName,
  qryClubInvoice.ClubContact,
  qryClubInvoice.ClubAddress1,
  qryClubInvoice.ClubAddress2,
  qryClubInvoice.ClubCityST,
  qryClubInvoice.ClubPhone1,
  qryClubInvoice.LineID,
  qryClubInvoice.ItemID,
  qryClubInvoice.Description,
  qryClubInvoice.SizeDesc,
  qryClubInvoice.QuantityShipped,
  qryClubInvoice.NetPrice,
  [QuantityShipped] * [NetPrice] AS XPrice,
  [QuantityShipped] * [Cost] AS XCost,
  qryClubInvoice.Cost,
  qryClubInvoice.OrderNotes,
  qryClubInvoice.SalePrice,
  qryClubInvoice.PricePerOZ,
  qryClubInvoice.ClubFax,
  qryClubInvoice.SalesDepartmentID,
  qryClubInvoice.CatDesc,
  qryClubInvoice.Location
FROM
  qryClubInvoice,
  Sysfile
ORDER BY
  qryClubInvoice.LineID;
