TRANSFORM Sum(ClubHistoryByBottleByMonth.Qty) AS SumOfQty
SELECT
  ClubHistoryByBottleByMonth.ClubID,
  ClubHistoryByBottleByMonth.ClubName,
  ClubHistoryByBottleByMonth.CatDesc,
  ClubHistoryByBottleByMonth.Code,
  ClubHistoryByBottleByMonth.Description,
  ClubHistoryByBottleByMonth.SizeD,
  Sum(ClubHistoryByBottleByMonth.Qty) AS TotalQty,
  Sum(
    ClubHistoryByBottleByMonth.Dollars
  ) AS TotalDollars
FROM
  ClubHistoryByBottleByMonth
GROUP BY
  ClubHistoryByBottleByMonth.ClubID,
  ClubHistoryByBottleByMonth.ClubName,
  ClubHistoryByBottleByMonth.CatDesc,
  ClubHistoryByBottleByMonth.Code,
  ClubHistoryByBottleByMonth.Description,
  ClubHistoryByBottleByMonth.SizeD
ORDER BY
  ClubHistoryByBottleByMonth.ClubName,
  ClubHistoryByBottleByMonth.CatDesc,
  ClubHistoryByBottleByMonth.Code PIVOT ClubHistoryByBottleByMonth.[Year-Month];
