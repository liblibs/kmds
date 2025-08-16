TRANSFORM Sum(ClubHistoryByBottleByMonth.Qty) AS SumOfQty
SELECT
  ClubHistoryByBottleByMonth.CatDesc,
  ClubHistoryByBottleByMonth.Code,
  ClubHistoryByBottleByMonth.Description,
  ClubHistoryByBottleByMonth.ClubID,
  ClubHistoryByBottleByMonth.ClubName,
  ClubHistoryByBottleByMonth.SizeD,
  Sum(ClubHistoryByBottleByMonth.Qty) AS TotalQty,
  Sum(
    ClubHistoryByBottleByMonth.Dollars
  ) AS TotalDollars
FROM
  ClubHistoryByBottleByMonth
GROUP BY
  ClubHistoryByBottleByMonth.CatDesc,
  ClubHistoryByBottleByMonth.Code,
  ClubHistoryByBottleByMonth.Description,
  ClubHistoryByBottleByMonth.ClubID,
  ClubHistoryByBottleByMonth.ClubName,
  ClubHistoryByBottleByMonth.SizeD
ORDER BY
  ClubHistoryByBottleByMonth.CatDesc,
  ClubHistoryByBottleByMonth.Code,
  ClubHistoryByBottleByMonth.ClubName PIVOT ClubHistoryByBottleByMonth.[Year-Month];
