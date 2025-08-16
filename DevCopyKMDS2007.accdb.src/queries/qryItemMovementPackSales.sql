SELECT
  Item.ID AS CartonID,
  Item.Description,
  UPCItem.ItemID AS PackID,
  IIf(
    IsNull([UnitsPerPack])
    Or [UnitsPerPack] = 0,
    1,
    [UnitsPerPack]
  ) AS PacksPerCarton,
  Sum(InvoLine.Quantity) AS PacksSold,
  Sum(InvoLine.Price) AS PackDollars,
  CDate(
    Format(
      InvoLine.SetupDate, "mm/dd/yyyy"
    )
  ) AS MMYY,
  Format(InvoLine.SetupDate, "mm") AS MM,
  Format(InvoLine.SetupDate, "yyyy") AS YY
FROM
  (
    Item
    INNER JOIN UPCItem ON Item.SubPack = UPCItem.ID
  )
  INNER JOIN InvoLine ON UPCItem.ItemID = InvoLine.Code
WHERE
  (
    (
      (InvoLine.TransType)= "S"
    )
  )
GROUP BY
  Item.ID,
  Item.Description,
  UPCItem.ItemID,
  IIf(
    IsNull([UnitsPerPack])
    Or [UnitsPerPack] = 0,
    1,
    [UnitsPerPack]
  ),
  CDate(
    Format(
      InvoLine.SetupDate, "mm/dd/yyyy"
    )
  ),
  Format(InvoLine.SetupDate, "mm"),
  Format(InvoLine.SetupDate, "yyyy");
