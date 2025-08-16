SELECT
  qryTodaysLiquorTransactions.Today,
  [CurrentLiquorValue] - [Counter] - [Clubs] - [Purchases] - [TransfersIn] - [TransfersOut] - [Adjustments] AS LastNight,
  qryTodaysLiquorTransactions.Counter,
  qryTodaysLiquorTransactions.Clubs,
  qryTodaysLiquorTransactions.Purchases,
  qryTodaysLiquorTransactions.TransfersOUT,
  qryTodaysLiquorTransactions.TransfersIN,
  qryTodaysLiquorTransactions.Adjustments,
  qryTodaysLiquorValueNow.CurrentLiquorValue
FROM
  qryTodaysLiquorTransactions,
  qryTodaysLiquorValueNow;
