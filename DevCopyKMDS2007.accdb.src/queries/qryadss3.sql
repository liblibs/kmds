INSERT INTO DaysSupplyADSForNewItems (
  Code, SumOfQuantity, SumOfPrice, MinOfSetupDate,
  MaxOfSetupDate, Days, ADS
)
SELECT
  qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code,
  qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfQuantity,
  qryDaysSupplyInvoLineItemSalesByMinMaxDate.SumOfPrice,
  qryDaysSupplyInvoLineItemSalesByMinMaxDate.MinOfSetupDate,
  qryDaysSupplyInvoLineItemSalesByMinMaxDate.MaxOfSetupDate,
  CInt(
    [MaxOfSetupDate] - [MinOfSetupDate]
  ) AS Days,
  (
    (
      [SumOfQuantity] / IIf([Days] < 30, 30, [days])
    )/ [UnitsPerPack]
  ) AS ADS
FROM
  (
    (
      Item
      INNER JOIN UPCItem ON Item.SubPack = UPCItem.ID
    )
    INNER JOIN qryDaysSupplyInvoLineItemSalesByMinMaxDate ON UPCItem.ItemID = qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code
  )
  LEFT JOIN DaysSupplyADSForNewItems ON qryDaysSupplyInvoLineItemSalesByMinMaxDate.Code = DaysSupplyADSForNewItems.Code
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
