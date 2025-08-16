SELECT
  qryFintechDivisionID.Division_ID,
  qryFintechTransactionClubID.Invoice_Number,
  Format(
    [qryFintechTransactionClubID].[SetupDate],
    "mm/dd/yyyy"
  ) AS invoice_date,
  qryFintechTransactionClubID.Licensee AS Vendor_store_id,
  InvoLine5.Quantity AS quantity_shipped,
  "EA" AS Quantity_uom,
  CStr([Item].[id]) AS item_number,
  IIf(
    ([involine5].[TransType] = "C"),
    [item].[Description],
    "Discount"
  ) AS product_description,
  [involine5].[price] / [involine5].[quantity] AS unit_price,
  Item.SalesDepartmentID
FROM
  Sysfile,
  qryFintechDivisionID,
  (
    (
      qryFinTechTransactions
      INNER JOIN qryFintechTransactionClubID ON (
        qryFinTechTransactions.SetupDate = qryFintechTransactionClubID.SetupDate
      )
      AND (
        qryFinTechTransactions.CashSysID = qryFintechTransactionClubID.CashSysID
      )
      AND (
        qryFinTechTransactions.InvoHeadID = qryFintechTransactionClubID.InvoHeadID
      )
    )
    INNER JOIN InvoLine5 ON (
      qryFinTechTransactions.SetupDate = InvoLine5.SetupDate
    )
    AND (
      qryFinTechTransactions.CashSysID = InvoLine5.CashSysID
    )
    AND (
      qryFinTechTransactions.InvoHeadID = InvoLine5.InvoHeadID
    )
  )
  INNER JOIN Item ON InvoLine5.Code = Item.ID
WHERE
  (
    (
      (InvoLine5.TransType)= "C"
      Or (InvoLine5.TransType)= "D"
    )
    AND (
      (Item.SalesDepartmentID)= 1
      Or (Item.SalesDepartmentID)= 2
    )
  );
