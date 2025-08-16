INSERT INTO Item (
  ID, SalesDepartmentID, VendorID, Description,
  Price, SetupDate, ChangeDate, SizeID,
  UnitsPerCase, UnitsPerOrder, OrderTo
)
SELECT
  InvoLine4.Code AS ID,
  1 AS SalesDepartmentID,
  1 AS VendorID,
  "= Setup from EOD" AS Description,
  InvoLine4.Price,
  Now() AS SetupDate,
  Now() AS ChangeDate,
  IIf(
    IsNull(Size.id),
    0,
    size.id
  ) AS SizeID,
  IIf(
    IsNull(Size.CaseQty),
    0,
    size.CaseQty
  ) AS UnitsPerCase,
  1 AS UnitsPerOrder,
  1 AS OrderTo
FROM
  (
    InvoLine4
    LEFT JOIN Item ON InvoLine4.Code = Item.ID
  )
  LEFT JOIN [Size] ON InvoLine4.Drawer = Size.ID
WHERE
  (
    (
      (InvoLine4.TransType)= "S"
    )
    AND (
      (Item.ID) Is Null
    )
  );
