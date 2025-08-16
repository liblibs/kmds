SELECT
  BuyDownItem.ItemsCode,
  BuyDownItem.SpcPrice,
  BuyDown.BeginDate,
  BuyDown.EndDate
FROM
  BuyDownItem
  INNER JOIN BuyDown ON BuyDownItem.BuyDownID = BuyDown.ID
WHERE
  (
    (
      (BuyDown.BeginDate)< Now()
    )
    AND (
      (BuyDown.EndDate)> Now()
    )
  );
