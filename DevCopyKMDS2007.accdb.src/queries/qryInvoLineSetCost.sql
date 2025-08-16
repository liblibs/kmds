UPDATE
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
SET
  InvoLine.Cost = Item.Cost;
