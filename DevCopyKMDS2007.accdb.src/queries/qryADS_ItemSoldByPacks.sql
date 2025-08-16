SELECT
  Item.ID,
  Item.UnitsPerPack,
  Sum(InvoLine.Quantity) AS SumOfQuantity
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
      Or (InvoLine.TransType)= "C"
    )
    AND (
      (InvoLine.SetupDate)> Now()-365
    )
  )
GROUP BY
  Item.ID,
  Item.UnitsPerPack
HAVING
  (
    (
      (Item.UnitsPerPack)> 0
    )
  );
