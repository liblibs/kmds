SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description,
  UPCItem.ItemID,
  Item.Description,
  UPCItem.ID,
  UPCItem.SetupDate,
  Len(UPCItem.id) AS lid
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
  )
ORDER BY
  Item.SalesDepartmentID,
  UPCItem.ItemID;
