INSERT INTO ClubInvoiceForCSVExport (
  AgentNumber, Name, ClubID, ClubName,
  InvoiceNo, InvDate, ItemID, Description,
  [Size], Quantity, Price, XPrice, SalePrice,
  PricePerOZ
)
SELECT
  First(Sysfile.AgentNumber) AS FirstOfAgentNumber,
  First(Sysfile.Name) AS FirstOfName,
  First(qryClubInvoice.ClubID) AS FirstOfClubID,
  First(qryClubInvoice.ClubName) AS FirstOfClubName,
  First(qryClubInvoice.InvoiceNo) AS FirstOfInvoiceNo,
  First(qryClubInvoice.InvDate) AS FirstOfInvDate,
  First(Sysfile.ClubDiscntItemID) AS FirstOfClubDiscntItemID,
  First("LiquorDiscount") AS des,
  First("LiqDisc") AS [Size],
  First(1) AS Quantity,
  First(0 - [LiquorDiscount]) AS Price,
  First(0 - [LiquorDiscount]) AS XPrice,
  First(qryClubInvoice.SalePrice) AS FirstOfSalePrice,
  First(0) AS PricePerOZ
FROM
  qryClubInvoice,
  Sysfile;
