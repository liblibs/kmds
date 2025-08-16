SELECT
  OrderLine.OrdheadID,
  Sum(
    IIf(
      SalesDepartment.ID = 1, Orderline.QuantityShipped,
      0
    )
  ) AS LiquorQuantity,
  Sum(
    IIf(
      SalesDepartment.ID = 1, Orderline.Price * [QuantityShipped],
      0
    )
  ) AS LiquorExtPrice,
  IIf(
    Club.ApplyDiscount, sysfile.ClubDiscRate,
    0
  ) AS LiquorDiscountRate,
  (
    (
      (
        [LiquorExtPrice] * [LiquorDiscountRate]
      ) \1
    )/ 100
  ) AS LiquorDiscount,
  Sum(
    IIf(
      SalesDepartment.ID = 1, Orderline.Cost * [QuantityShipped],
      0
    )
  ) AS LiquorExtCost,
  Club.NonLiquor,
  Sum(
    IIf(
      SalesDepartment.ID <> 1, Orderline.QuantityShipped,
      0
    )
  ) AS NonLiquorQuantity,
  Sum(
    IIf(
      SalesDepartment.ID <> 1,
      (
        Orderline.Price - OrderLine.BuyDown
      )* [QuantityShipped],
      0
    )
  ) AS NonLiquorExtPrice,
  IIf(
    Club.ApplyDiscount, sysfile.NonClubDiscRate,
    0
  ) AS NonLiquorDiscountRate,
  Sum(
    IIf(
      SalesDepartment.ID <> 1,
      IIf(
        SalesDepartment.AllowClubDiscounts,
        (
          Orderline.Price - OrderLine.BuyDown
        )* Orderline.QuantityShipped,
        0
      ),
      0
    )
  ) AS NonLiquorExtPriceDiscountable,
  (
    (
      (
        [NonLiquorExtPriceDiscountable] * [NonLiquorDiscountRate]
      ) \1
    )/ 100
  ) AS NonLiquorDiscount,
  Sum(
    IIf(
      SalesDepartment.ID <> 1, Orderline.Cost * [QuantityShipped],
      0
    )
  ) AS NonLiquorExtCost,
  Sum(OrderLine.QuantityShipped) AS Quantity,
  [LiquorExtPrice] + [NonLiquorExtPrice] AS ExtPrice,
  [LiquorExtCost] + [NonLiquorExtCost] AS ExtCost
FROM
  Sysfile,
  (
    (
      (
        Club
        INNER JOIN OrdHead ON Club.ID = OrdHead.ClubID
      )
      INNER JOIN OrderLine ON OrdHead.ID = OrderLine.OrdheadID
    )
    INNER JOIN Item ON OrderLine.ItemID = Item.ID
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
GROUP BY
  OrderLine.OrdheadID,
  IIf(
    Club.ApplyDiscount, sysfile.ClubDiscRate,
    0
  ),
  Club.NonLiquor,
  IIf(
    Club.ApplyDiscount, sysfile.NonClubDiscRate,
    0
  );
