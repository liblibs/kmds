PARAMETERS parBeginDate DateTime,
parEndDate DateTime,
parVendorID Long,
parBeginSalesDepartmentID Long,
parEndSalesDepartmentID Long;
SELECT
  [parBeginDate] AS DateStart,
  [parEndDate] AS DateEnd,
  Item.VendorID,
  InvoLine.Code AS ItemId,
  Sum(
    IIf(
      [TransType] = "X", 0 - [Quantity], [Quantity]
    )
  ) AS Sold,
  (
    DateDiff(
      "d", [parBeginDate], [parEndDate]
    )+ 1
  ) AS SoldDays
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.SetupDate) Between [parBeginDate]
      And [parEndDate]
    )
    AND (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "X"
    )
    AND (
      (Item.SalesDepartmentID) Between [parBeginSalesDepartmentID]
      And [parEndSalesDepartmentID]
    )
  )
GROUP BY
  Item.VendorID,
  InvoLine.Code
HAVING
  (
    (
      (Item.VendorID)= [parVendorId]
    )
  )
ORDER BY
  InvoLine.Code;
