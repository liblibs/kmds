SELECT
  qryOLCCDepositRecapSales.Counter,
  qryOLCCDepositRecapSales.ClubsGross,
  qryOLCCDepositRecapSales.ClubsDiscount,
  qryOLCCDepositRecapSales.ClubsNet,
  qryOLCCDepositRecapSales.OLCCDeposit,
  qryOLCCDepositRecapTenders.Checks,
  qryOLCCDepositRecapTenders.Cards,
  [OLCCDeposit] - [Checks] - [Cards] - [EFT] AS CashRequired,
  [OLCCDeposit] - [Cards] - [EFT] AS Deposit,
  qryOLCCDepositRecapTenders.EFT
FROM
  qryOLCCDepositRecapSales,
  qryOLCCDepositRecapTenders;
