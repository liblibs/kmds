SELECT
  UPCItem.ID
FROM
  (
    UPCItem
    INNER JOIN Item ON UPCItem.ItemID = Item.ID
  )
  INNER JOIN vwItemsWithCurrentPricing ON Item.ID = vwItemsWithCurrentPricing.ID
WHERE
  (
    (
      (Item.SalesDepartmentID) Between Forms!PSC!cboSalesDepartmentIDBegin
      And Forms!PSC!cboSalesDepartmentIDEnd
    )
    And (
      (
        Len(UPCItem.ID)
      )= 4
      Or (
        Len(UPCItem.ID)
      )= 5
      Or (
        Len(UPCItem.ID)
      )= 6
      Or (
        Len(UPCItem.ID)
      )= 8
      Or (
        Len(UPCItem.ID)
      )= 12
    )
  )
ORDER BY
  UPCItem.ID;
