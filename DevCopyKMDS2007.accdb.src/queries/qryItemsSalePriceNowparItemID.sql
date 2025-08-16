PARAMETERS parItemID Decimal;
SELECT
  qryItemsSalePriceNow.ItemID,
  qryItemsSalePriceNow.StartDate,
  qryItemsSalePriceNow.EndDate,
  qryItemsSalePriceNow.SalePrice
FROM
  qryItemsSalePriceNow
WHERE
  (
    (
      (qryItemsSalePriceNow.ItemID)= [parItemID]
    )
  );
