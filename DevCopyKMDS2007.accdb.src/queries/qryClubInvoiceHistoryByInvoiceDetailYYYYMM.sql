SELECT
  Club.NonLiquor,
  qryClubInvoiceHistoryByInvoice.ClubID,
  Club.Name AS ClubName,
  Format(
    qryClubInvoiceHistoryByInvoice.SETUPDATE,
    "yyyy-mm"
  ) AS YYYYMM,
  Sum(
    IIf(
      InvoLine.TransType = 'C', [Quantity],
      0
    )
  ) AS Qty,
  Sum(
    IIf(
      InvoLine.TransType = 'C', InvoLine.Price,
      0
    )
  ) AS Gross,
  Sum(
    IIf(
      InvoLine.TransType = 'D', InvoLine.Price,
      0
    )
  ) AS Discount,
  Sum(
    IIf(
      Item.SalesDepartmentID = 1
      And InvoLine.TransType = 'C',
      InvoLine.Price, 0
    )
  ) AS GrossLiquor,
  Sum(
    IIf(
      InvoLine.TransType = 'D'
      And InvoLine.Code = Sysfile.ClubDiscntItemID,
      InvoLine.Price, 0
    )
  ) AS LiquorDiscount,
  Sum(
    IIf(
      Item.SalesDepartmentID <> 1
      And InvoLine.TransType = 'C',
      InvoLine.Price, 0
    )
  ) AS GrossNonLiquor,
  Sum(
    IIf(
      InvoLine.TransType = 'D'
      And InvoLine.Code = Sysfile.NonClubDiscntItemID,
      InvoLine.Price, 0
    )
  ) AS NonLiquorDiscount
FROM
  Sysfile,
  (
    (
      qryClubInvoiceHistoryByInvoice
      INNER JOIN InvoLine ON (
        qryClubInvoiceHistoryByInvoice.SetupDate = InvoLine.SetupDate
      )
      AND (
        qryClubInvoiceHistoryByInvoice.CashSysID = InvoLine.CashSysID
      )
      AND (
        qryClubInvoiceHistoryByInvoice.InvoHeadID = InvoLine.InvoHeadID
      )
    )
    INNER JOIN Club ON qryClubInvoiceHistoryByInvoice.ClubID = Club.ID
  )
  INNER JOIN Item ON InvoLine.Code = Item.ID
GROUP BY
  Club.NonLiquor,
  qryClubInvoiceHistoryByInvoice.ClubID,
  Club.Name,
  Format(
    qryClubInvoiceHistoryByInvoice.SETUPDATE,
    "yyyy-mm"
  )
ORDER BY
  Club.Name,
  Format(
    qryClubInvoiceHistoryByInvoice.SETUPDATE,
    "yyyy-mm"
  );
