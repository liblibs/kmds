SELECT
  Size.Description,
  Sum(1) AS [Count],
  First(Size.CaseQty) AS CQ,
  Sum(OrderLine.QuantityShipped) AS QS
FROM
  OrdHead
  INNER JOIN (
    (
      Item
      LEFT JOIN [Size] ON Item.SizeID = Size.ID
    )
    INNER JOIN OrderLine ON Item.ID = OrderLine.ItemID
  ) ON OrdHead.ID = OrderLine.OrdheadID
WHERE
  (
    (
      (OrdHead.InvoiceNo)= Forms!ClubOrderPrintDialog!txtInvoiceNo
    )
    And (
      (OrderLine.QuantityShipped)<> 0
    )
    And (
      (Item.SalesDepartmentID)= 1
    )
  )
GROUP BY
  Size.Description;
