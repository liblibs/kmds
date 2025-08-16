SELECT
  Club.ID,
  Club.Name,
  Item.ID,
  Item.Description,
  Sum(
    qryInvoLineClubSalesByItemByWeek.SumOfQuantity
  ) AS SumOfSumOfQuantity,
  Sum(
    IIf([week] = 1, [SumOfQuantity], 0)
  ) AS Week1,
  Sum(
    IIf([week] = 2, [SumOfQuantity], 0)
  ) AS Week2,
  Sum(
    IIf([week] = 3, [SumOfQuantity], 0)
  ) AS Week3,
  Sum(
    IIf([week] = 4, [SumOfQuantity], 0)
  ) AS Week4,
  Sum(
    IIf([week] > 4, [SumOfQuantity], 0)
  ) AS Week5
FROM
  (
    qryInvoLineClubSalesByItemByWeek
    INNER JOIN Club ON qryInvoLineClubSalesByItemByWeek.ClubId = Club.ID
  )
  INNER JOIN Item ON qryInvoLineClubSalesByItemByWeek.Code = Item.ID
GROUP BY
  Club.ID,
  Club.Name,
  Item.ID,
  Item.Description;
