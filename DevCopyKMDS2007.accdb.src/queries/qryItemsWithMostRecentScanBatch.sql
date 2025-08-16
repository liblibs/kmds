SELECT
  IIf(
    IsNull([MaxOfSetupDate]),
    0,
    [MaxOfSetupDate]
  ) AS MostRecentCount,
  vwItemsWithCurrentPricing.SalesDepartmentID,
  vwItemsWithCurrentPricing.SalesDptDescr,
  vwItemsWithCurrentPricing.SubCategoryID,
  vwItemsWithCurrentPricing.CategoryDescr,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.Onhand,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.Cost
FROM
  vwItemsWithCurrentPricing
  LEFT JOIN qryInvoLineYTransMostRecent ON vwItemsWithCurrentPricing.ID = qryInvoLineYTransMostRecent.Code
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
        vwItemsWithCurrentPricing.CategoryDescr
      ) Between Forms!PrintInventoryReport!cboCategoryIDBegin
      And Forms!PrintInventoryReport!cboCategoryIDEnd
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
        vwItemsWithCurrentPricing.CategoryDescr
      ) Between Forms!PrintInventoryReport!cboCategoryIDBegin
      And Forms!PrintInventoryReport!cboCategoryIDEnd
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
