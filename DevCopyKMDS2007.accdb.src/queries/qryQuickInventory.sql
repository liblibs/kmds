SELECT
  vwItemsWithCurrentPricing.SalesDepartmentID,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Onhand,
  IIf(
    [OnShelf] <> [Onhand], [Onshelf], 0
  ) AS OnTheShelf,
  vwItemsWithCurrentPricing.Description
FROM
  vwItemsWithCurrentPricing
WHERE
  (
    (
      (
        vwItemsWithCurrentPricing.SalesDepartmentID
      ) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (
        vwItemsWithCurrentPricing.Onhand
      )<> 0
    )
    And (
      (
        IIf(
          Forms!PrintInventoryReport!chkPrintZeroOnhands,
          -1, 0
        )
      )= 0
    )
  )
  Or (
    (
      (
        vwItemsWithCurrentPricing.SalesDepartmentID
      ) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin
      And Forms!PrintInventoryReport!cboSalesDepartmentIDEnd
    )
    And (
      (
        IIf(
          Forms!PrintInventoryReport!chkPrintZeroOnhands,
          -1, 0
        )
      )=-1
    )
  );
