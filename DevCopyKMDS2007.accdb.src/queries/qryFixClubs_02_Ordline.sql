INSERT INTO OrderLine (
  OrdheadID, LineID, ItemID, QuantityOrdered,
  QuantityShipped, Price, Cost, SetupDate,
  ChangeDate
)
SELECT
  OrdHead.ID AS OrdheadID,
  InvoLine.ID AS LineID,
  InvoLine.Code AS ItemID,
  InvoLine.Quantity,
  InvoLine.Quantity,
  [InvoLine].[Price] / [InvoLine].[Quantity] AS UP,
  [InvoLine].[Price] / [InvoLine].[Quantity] AS UT,
  InvoLine.SetupDate,
  InvoLine.SetupDate
FROM
  OrdHead
  INNER JOIN (
    qryFixClubOrders
    INNER JOIN InvoLine ON (
      qryFixClubOrders.SetupDate = InvoLine.SetupDate
    )
    AND (
      qryFixClubOrders.CashSysID = InvoLine.CashSysID
    )
    AND (
      qryFixClubOrders.InvoHeadID = InvoLine.InvoHeadID
    )
  ) ON OrdHead.InvoiceNo = qryFixClubOrders.InvoiceNumber
WHERE
  (
    (
      (InvoLine.TransType)= "C"
    )
  )
ORDER BY
  OrdHead.ID,
  InvoLine.ID;
