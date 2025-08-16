SELECT
  Sysfile.*,
  qryClubInvoicePOSPrinter.ClubID,
  qryClubInvoicePOSPrinter.Order,
  qryClubInvoicePOSPrinter.InvoiceNo,
  qryClubInvoicePOSPrinter.OrdDate,
  qryClubInvoicePOSPrinter.InvDate,
  qryClubInvoicePOSPrinter.ClubName,
  qryClubInvoicePOSPrinter.ClubContact,
  qryClubInvoicePOSPrinter.ClubAddress1,
  qryClubInvoicePOSPrinter.ClubAddress2,
  qryClubInvoicePOSPrinter.ClubCityST,
  qryClubInvoicePOSPrinter.ClubPhone1,
  qryClubInvoicePOSPrinter.LineID,
  qryClubInvoicePOSPrinter.ItemID,
  qryClubInvoicePOSPrinter.Description,
  qryClubInvoicePOSPrinter.SizeDesc,
  qryClubInvoicePOSPrinter.QuantityShipped,
  qryClubInvoicePOSPrinter.Price,
  [QuantityShipped] * [Price] AS XPrice,
  [QuantityShipped] * [Cost] AS XCost,
  qryClubInvoicePOSPrinter.Cost,
  qryClubInvoicePOSPrinter.OrderNotes,
  qryClubInvoicePOSPrinter.SalePrice,
  qryClubInvoicePOSPrinter.PricePerOZ
FROM
  Sysfile,
  qryClubInvoicePOSPrinter
ORDER BY
  qryClubInvoicePOSPrinter.ItemID;
