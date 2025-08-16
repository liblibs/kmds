SELECT
  Sysfile.AgentNumber,
  Sysfile.Name,
  qryClubInvoice.ClubID,
  qryClubInvoice.ClubName,
  qryClubInvoice.InvoiceNo,
  qryClubInvoice.InvDate,
  qryClubInvoice.ItemID,
  qryClubInvoice.Description,
  qryClubInvoice.SizeDesc AS [Size],
  qryClubInvoice.QuantityShipped AS Quantity,
  qryClubInvoice.NetPrice AS Price,
  [QuantityShipped] * [NetPrice] AS XPrice,
  qryClubInvoice.SalePrice,
  qryClubInvoice.PricePerOZ INTO ClubInvoiceForCSVExport
FROM
  qryClubInvoice,
  Sysfile;
