SELECT
  SalesDepartment.Description,
  Item.ID,
  Item.Description,
  InvoLine.SetupDate,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.TransType,
  InvoLine.InputCode
FROM
  InvoLine
  INNER JOIN (
    Item
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  ) ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (SalesDepartment.Description) Like "RJ*"
    )
    AND (
      (InvoLine.SetupDate) Between #1/1/2007#
      And #1/1/2008#
    )
    AND (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
    )
  )
ORDER BY
  SalesDepartment.Description,
  Item.Description,
  InvoLine.SetupDate;
