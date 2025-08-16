SELECT
  Sysfile.AgentNumber,
  Sysfile.Name,
  Sysfile.Add1,
  Sysfile.Add2,
  Sysfile.Add3,
  Sysfile.Phone,
  [CashSysId] & "-" & [ShiftID] & "-" & [InvoHeadID] AS InvoiceNumber,
  InvoLine5.SetupDate,
  0 AS CustomerID,
  InvoLine5.Code AS SKU,
  Item.Description,
  Size.Description AS SizeDesc,
  IIf(
    [transtype] = "C", [Code], [InputCode]
  ) AS UPC,
  InvoLine5.Price / [quantity] AS UnitPrice,
  InvoLine5.Quantity,
  InvoLine5.Price AS ExtendPrice
FROM
  Sysfile,
  (
    InvoLine5
    INNER JOIN Item ON InvoLine5.Code = Item.ID
  )
  INNER JOIN [Size] ON InvoLine5.ID = Size.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  )
ORDER BY
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  InvoLine5.ID;
