UPDATE
  Sysfile,
  qryBottleDepositQtyAdjAtEOD,
  InvoLine5
  INNER JOIN Item ON InvoLine5.Code = Item.ID
SET
  InvoLine5.Quantity = [Involine5].[Price] / [item].[price],
  InvoLine5.Cost = [InvoLine5].[PRICE]
WHERE
  (
    (
      (InvoLine5.Code)= [DepositItemID]
    )
    AND (
      (Item.Price)> 0
    )
  );
