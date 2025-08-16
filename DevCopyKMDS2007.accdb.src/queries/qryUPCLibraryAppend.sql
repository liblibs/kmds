INSERT INTO UPCLibrary (UPCCode, Description)
SELECT
  UPCItem.ID,
  Item.Description
FROM
  (
    UPCItem
    INNER JOIN Item ON UPCItem.ItemID = Item.ID
  )
  LEFT JOIN UPCLibrary ON UPCItem.ID = UPCLibrary.UPCCode
WHERE
  (
    (
      (UPCLibrary.UPCCode) Is Null
    )
    And (
      (
        Len(UPCItem.ID)
      )> 7
    )
    And (
      (Item.SalesDepartmentID)> 1
    )
  );
