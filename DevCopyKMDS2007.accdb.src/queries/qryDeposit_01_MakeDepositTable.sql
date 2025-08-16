SELECT
  Sum([onhand] * [deposit]) AS DepositTotal,
  Sysfile.DepositItemID INTO DepositTable
FROM
  Item,
  Sysfile
GROUP BY
  Sysfile.DepositItemID;
