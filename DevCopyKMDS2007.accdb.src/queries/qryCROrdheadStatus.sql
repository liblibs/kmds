SELECT
  OrdHead.ID,
  InvoiceNo,
  Club.Name,
  OrdHead.InvDate,
  OrdHead.InvDollars,
  OrdHead.PaidFlag
FROM
  OrdHead
  INNER JOIN Club ON OrdHead.ClubID = Club.ID;
