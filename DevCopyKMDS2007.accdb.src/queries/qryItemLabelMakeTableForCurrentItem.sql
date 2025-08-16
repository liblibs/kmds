PARAMETERS parItemID Decimal;
SELECT
  Item.ID,
  Item.Description AS ItemD,
  IIf(
    size.Description = "Not Assigned",
    "", size.description
  ) AS SizeD,
  IIf(
    IsNull([SalePrice]),
    [Price],
    [SalePrice]
  ) AS CurPrice,
  IIf(
    IsNull([SalePrice]),
    "",
    "*"
  ) AS S,
  SalesDepartment.Description AS SalesDpt,
  Item.Price,
  Vendor.Name AS VendorName,
  qryUPCItemMax.UPCCodeMax,
  Item.SetupDate,
  Item.ChangeDate,
  1 AS Labels INTO ItemLabels
FROM
  (
    (
      (
        (
          Item
          INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
        )
        LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
      )
      INNER JOIN [Size] ON Item.SizeID = Size.ID
    )
    LEFT JOIN Vendor ON Item.VendorID = Vendor.ID
  )
  LEFT JOIN qryUPCItemMax ON Item.ID = qryUPCItemMax.ItemID
WHERE
  (
    (
      (Item.ID)= [parItemID]
    )
  );
