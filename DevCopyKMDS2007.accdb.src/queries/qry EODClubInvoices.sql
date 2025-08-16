SELECT
  Club.NonLiquor,
  [qry EOD060].ClubID,
  Club.Name AS ClubName,
  InvoLine5.InvoHeadID,
  Sum(
    IIf(
      InvoLine5.TransType = 'C', [Quantity],
      0
    )
  ) AS Qty,
  Sum(
    IIf(
      InvoLine5.TransType = 'C', InvoLine5.Price,
      0
    )
  ) AS Gross,
  Sum(
    IIf(
      InvoLine5.TransType = 'D', InvoLine5.Price,
      0
    )
  ) AS Discount,
  Sum(
    IIf(
      Item.SalesDepartmentID = 1
      And InvoLine5.TransType = 'C',
      involine5.Price, 0
    )
  ) AS GrossLiquor,
  Sum(
    IIf(
      InvoLine5.TransType = 'D'
      And Involine5.Code = Sysfile.ClubDiscntItemID,
      InvoLine5.Price, 0
    )
  ) AS LiquorDiscount,
  Sum(
    IIf(
      Item.SalesDepartmentID <> 1
      And InvoLine5.TransType = 'C',
      involine5.Price, 0
    )
  ) AS GrossNonLiquor,
  Sum(
    IIf(
      InvoLine5.TransType = 'D'
      And Involine5.Code = Sysfile.NonClubDiscntItemID,
      InvoLine5.Price, 0
    )
  ) AS NonLiquorDiscount,
  [qry EOD060InvoiceNumber].InvoiceNumber
FROM
  Sysfile,
  (
    (
      (
        [qry EOD060]
        INNER JOIN InvoLine5 ON [qry EOD060].InvoHeadID = InvoLine5.InvoHeadID
      )
      INNER JOIN Club ON [qry EOD060].ClubID = Club.ID
    )
    INNER JOIN Item ON InvoLine5.Code = Item.ID
  )
  INNER JOIN [qry EOD060InvoiceNumber] ON (
    InvoLine5.InvoHeadID = [qry EOD060InvoiceNumber].InvoHeadID
  )
  AND (
    InvoLine5.CashSysID = [qry EOD060InvoiceNumber].CashSysID
  )
GROUP BY
  Club.NonLiquor,
  [qry EOD060].ClubID,
  Club.Name,
  InvoLine5.InvoHeadID,
  [qry EOD060InvoiceNumber].InvoiceNumber;
