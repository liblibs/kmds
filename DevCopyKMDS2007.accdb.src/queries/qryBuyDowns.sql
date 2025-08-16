SELECT
  BuyDown.ID,
  BuyDown.RepsName,
  BuyDown.ReferenceNo,
  BuyDown.BeginDate,
  BuyDown.EndDate,
  BuyDown.BuydownDollars,
  BuyDown.BuydownQty,
  BuyDown.ItemsID AS PromoItemID,
  Item.Description AS PromItemDesc,
  BuyDownItem.ItemsCode,
  Item_1.Description,
  BuyDownItem.QuantitySold,
  BuyDownItem.DiscountDollars,
  BuyDownItem.ShpQty,
  BuyDownItem.ShpDollars,
  SalesDepartment.Description AS SalesdptD,
  Size.Description AS SizeD
FROM
  (
    (
      (
        (
          BuyDown
          INNER JOIN BuyDownItem ON BuyDown.ID = BuyDownItem.BuyDownID
        )
        INNER JOIN Item ON BuyDown.ItemsID = Item.ID
      )
      INNER JOIN Item AS Item_1 ON BuyDownItem.ItemsCode = Item_1.ID
    )
    INNER JOIN [Size] ON Item_1.SizeID = Size.ID
  )
  INNER JOIN SalesDepartment ON Item_1.SalesDepartmentID = SalesDepartment.ID;
