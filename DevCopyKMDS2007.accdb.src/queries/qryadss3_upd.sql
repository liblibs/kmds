UPDATE
  DaysSupplyADSForNewItems
  INNER JOIN (
    (
      Item
      INNER JOIN UPCItem ON Item.SubPack = UPCItem.ID
    )
    INNER JOIN qryDaysSupplyInvoLineItemSalesByMinMaxDate ON UPCItem.ItemID = qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code
  ) ON DaysSupplyADSForNewItems.Code = qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code
SET
  DaysSupplyADSForNewItems.SumOfQuantity = DaysSupplyADSForNewItems.sumofquantity +(
    (
      qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfQuantity / IIf(
        qryDaysSupplyInvoLineItemSalesByMinMaxDate.Days < 30,
        30, qryDaysSupplyInvoLineItemSalesByMinMaxDate.days
      )
    )/ [UnitsPerPack]
  )
WHERE
  (
    (
      (
        qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate
      )> Now()-365
    )
    AND (
      (
        qryDaysSupplyInvoLineItemSalesByMinMaxDate.MaxOfSetupDate
      )> Now()-90
    )
    AND (
      (DaysSupplyADSForNewItems.Code) Is Null
    )
  );
