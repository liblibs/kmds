INSERT INTO OrdHead (
  ClubID, [Order], InvoiceNo, SetupDate,
  InvDate, OrdDate, ChangeDate, DisDollars,
  LiquorDiscount
)
SELECT
  qryFixClubOrders.ClubID,
  qryFixClubOrders.InvoiceNumber,
  qryFixClubOrders.InvoiceNumber,
  qryFixClubOrders.SetupDate,
  qryFixClubOrders.SetupDate,
  qryFixClubOrders.SetupDate,
  qryFixClubOrders.SetupDate,
  0 - [Price] AS p,
  0 - [Price] AS q
FROM
  qryFixClubOrders;
