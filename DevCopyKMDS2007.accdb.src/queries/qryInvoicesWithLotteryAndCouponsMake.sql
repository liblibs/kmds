SELECT
  qryInvoicesWithCoupons.InvoHeadID,
  qryInvoicesWithCoupons.CashSysID,
  qryInvoicesWithCoupons.SetupDate INTO InvoLineWithLotteryAndCoupons
FROM
  qryInvoicesWithCoupons
  INNER JOIN qryInvoicesWithLottery ON (
    qryInvoicesWithCoupons.SetupDate = qryInvoicesWithLottery.SetupDate
  )
  AND (
    qryInvoicesWithCoupons.CashSysID = qryInvoicesWithLottery.CashSysID
  )
  AND (
    qryInvoicesWithCoupons.InvoHeadID = qryInvoicesWithLottery.InvoHeadID
  );
