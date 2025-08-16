DELETE UPCItem.SetupDate,
InvoLine.InvoHeadID,
Len(UPCItem.id) AS Expr1,
UPCItem.*,
Len([UPCItem].[id]) AS Expr2
FROM
  (
    (
      UPCItem
      INNER JOIN Item ON UPCItem.ItemID = Item.ID
    )
    LEFT JOIN InvoLine ON UPCItem.ID = InvoLine.InputCode
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
WHERE
  (
    (
      (UPCItem.SetupDate)< Now()-60
    )
    And (
      (InvoLine.InvoHeadID) Is Null
    )
    And (
      (
        Len(UPCItem.id)
      )> Len([itemid])
    )
  );
