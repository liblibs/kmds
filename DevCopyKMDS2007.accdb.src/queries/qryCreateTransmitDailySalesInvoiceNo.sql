SELECT
  InvoLine5.InvoHeadID,
  InvoLine5.SetupDate,
  InvoLine5.CashSysID,
  InvoLine5.Code AS DiscountID,
  InvoLine5.Price AS Discount,
  Format([InputCode], "0000000000") AS LicenseeInvoiceNumber,
  Club.ID AS ClubID,
  Club.Name AS ClubName,
  Format([Licensee], "0000000000") AS PremisesNumber
FROM
  (
    Sysfile
    INNER JOIN InvoLine5 ON Sysfile.ClubDiscntItemID = InvoLine5.Code
  )
  INNER JOIN Club ON InvoLine5.SalesDept = Club.ID;
