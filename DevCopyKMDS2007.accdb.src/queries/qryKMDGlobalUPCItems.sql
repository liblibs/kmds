SELECT
  UPCItem.ID,
  UPCItem.ItemID,
  Sysfile.AgentNumber,
  Len([UPCItem].[id]) AS lid,
  UPCItem.ChangeDate
FROM
  Sysfile,
  UPCItem
  INNER JOIN Item ON UPCItem.ItemID = Item.ID
WHERE
  (
    (
      (
        Len(UPCItem.id)
      )> 5
    )
    And (
      (Item.SalesDepartmentID)= 1
    )
  );
