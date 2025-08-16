SELECT
  InvoLine5.CashSysID,
  Format(
    First(
      (
        (InvoLine5.InvoHeadID * 100)+ [ID]
      )
    ),
    "000000000"
  ) AS TransactionNumber,
  Format(
    Now(),
    'yymmdd'
  ) AS CreateDate,
  Format(
    Now(),
    'hhnn'
  ) AS CreateTime,
  Format([Code], "00000000000") AS OLCCItemCode,
  Format(
    Sum([quantity]),
    "00000;-0000"
  ) AS OLCCQuantity,
  Format(InvoLine5.SetupDate, 'yymmdd') AS RegisterDate,
  Format(InvoLine5.SetupDate, 'hhnnss') AS RegisterTime,
  Format(
    IIf(
      IsNumeric([InputCode]),
      [InputCode],
      [Code]
    ),
    "!@@@@@@@@@@@@@@"
  ) AS UPCCode,
  Format(
    (
      (InvoLine5.Price / [quantity])* 100
    ),
    "0000000"
  ) AS SellingPrice,
  qryCreateTransmitDailySalesInvoiceNo.LicenseeInvoiceNumber AS LIN,
  qryCreateTransmitDailySalesInvoiceNo.PremisesNumber AS PN,
  qryCreateTransmitDailySalesInvoiceNo.Discount,
  qryCreateTransmitDailySalesTenderCode.TenderCode
FROM
  (
    InvoLine5
    LEFT JOIN qryCreateTransmitDailySalesInvoiceNo ON (
      InvoLine5.InvoHeadID = qryCreateTransmitDailySalesInvoiceNo.InvoHeadID
    )
    AND (
      InvoLine5.CashSysID = qryCreateTransmitDailySalesInvoiceNo.CashSysID
    )
    AND (
      InvoLine5.SetupDate = qryCreateTransmitDailySalesInvoiceNo.SetupDate
    )
  )
  INNER JOIN qryCreateTransmitDailySalesTenderCode ON (
    InvoLine5.SetupDate = qryCreateTransmitDailySalesTenderCode.SetupDate
  )
  AND (
    InvoLine5.CashSysID = qryCreateTransmitDailySalesTenderCode.CashSysID
  )
  AND (
    InvoLine5.InvoHeadID = qryCreateTransmitDailySalesTenderCode.InvoHeadID
  )
GROUP BY
  InvoLine5.CashSysID,
  Format(
    Now(),
    'yymmdd'
  ),
  Format(
    Now(),
    'hhnn'
  ),
  Format([Code], "00000000000"),
  Format(InvoLine5.SetupDate, 'yymmdd'),
  Format(InvoLine5.SetupDate, 'hhnnss'),
  Format(
    IIf(
      IsNumeric([InputCode]),
      [InputCode],
      [Code]
    ),
    "!@@@@@@@@@@@@@@"
  ),
  Format(
    (
      (InvoLine5.Price / [quantity])* 100
    ),
    "0000000"
  ),
  qryCreateTransmitDailySalesInvoiceNo.LicenseeInvoiceNumber,
  qryCreateTransmitDailySalesInvoiceNo.PremisesNumber,
  qryCreateTransmitDailySalesInvoiceNo.Discount,
  qryCreateTransmitDailySalesTenderCode.TenderCode,
  InvoLine5.TransType
HAVING
  (
    (
      (
        Format([Code], "00000000000")
      ) Between 99900000000
      And 99999999999
    )
    AND (
      (
        Format(
          Sum([quantity]),
          "00000;-0000"
        )
      )<> 0
    )
    AND (
      (InvoLine5.TransType)= "S"
      Or (InvoLine5.TransType)= "C"
    )
  )
ORDER BY
  Format(
    First(
      (
        (InvoLine5.InvoHeadID * 100)+ [ID]
      )
    ),
    "000000000"
  );
