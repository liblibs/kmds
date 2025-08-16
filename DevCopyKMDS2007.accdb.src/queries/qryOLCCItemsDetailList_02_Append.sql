INSERT INTO OLCCItemsDetailList (
  OLCCCode, Desription, Category, [Size],
  Status, UnitsPerOrder, UnitsPerCase,
  PricePerBottle, QuantityAvailable
)
SELECT
  OLCCItemsDetailListXLS.F1,
  OLCCItemsDetailListXLS.F2,
  OLCCItemsDetailListXLS.F3,
  OLCCItemsDetailListXLS.F4,
  OLCCItemsDetailListXLS.F5,
  OLCCItemsDetailListXLS.F6,
  OLCCItemsDetailListXLS.F7,
  OLCCItemsDetailListXLS.F8,
  OLCCItemsDetailListXLS.F9
FROM
  OLCCItemsDetailListXLS;
