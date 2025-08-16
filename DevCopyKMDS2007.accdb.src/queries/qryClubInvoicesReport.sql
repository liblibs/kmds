SELECT
  Club.NonLiquor,
  qryClubInvoices.ClubID,
  Club.Name AS ClubName,
  InvoLineAllRegisters.InvoHeadID,
  Sum(
    IIf(
      InvoLineAllRegisters.TransType = 'C',
      [Quantity], 0
    )
  ) AS Qty,
  Sum(
    IIf(
      InvoLineAllRegisters.TransType = 'C',
      [Price], 0
    )
  ) AS Gross,
  Sum(
    IIf(
      InvoLineAllRegisters.TransType = 'D',
      [Price], 0
    )
  ) AS Discount
FROM
  (
    qryClubInvoices
    INNER JOIN InvoLineAllRegisters ON qryClubInvoices.InvoHeadID = InvoLineAllRegisters.InvoHeadID
  )
  INNER JOIN Club ON qryClubInvoices.ClubID = Club.ID
GROUP BY
  Club.NonLiquor,
  qryClubInvoices.ClubID,
  Club.Name,
  InvoLineAllRegisters.InvoHeadID;
