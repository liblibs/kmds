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
    [SumOfQuantity] / IIf([Days] < 30, 30, [days])
  ) AS ADS INTO DaysSupplyADSForNewItems
FROM
  qryDaysSupplyInvoLineItemSalesByMinMaxDate
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
  );
