SELECT
  Club.ID,
  Club.Name AS ClubName,
  Club.Contact AS ClubContact,
  Club.Address1 AS ClubAddress1,
  Club.Address2 AS ClubAddress2,
  [City] & ", " & [ST] & "  " & [Zip] AS ClubCityST,
  Club.Phone1 AS ClubPhone1,
  ClubOrderGuides.LineID,
  SubCategory.Description AS Category,
  ClubOrderGuides.ItemID,
  Item.Description,
  Size.Description AS SizeDesc,
  Size.Oz,
  CCur([price] / [oz]) AS PricePerOZ,
  qryItemsSalePriceNow.SalePrice,
  Item.Price,
  IIf(
    IsNull([SalePrice]),
    [Price],
    [SalePrice]
  ) AS TodaysPrice,
  IIf(
    IsNull([SalePrice]),
    "",
    "*"
  ) AS Sprice,
  Club.Phone2 AS ClubFax,
  ClubOrderGuides.Par
FROM
  (
    (
      (
        Item
        LEFT JOIN [Size] ON Item.SizeID = Size.ID
      )
      INNER JOIN (
        Club
        INNER JOIN ClubOrderGuides ON Club.ID = ClubOrderGuides.ClubID
      ) ON Item.ID = ClubOrderGuides.ItemID
    )
    LEFT JOIN qryItemsSalePriceNow ON Item.ID = qryItemsSalePriceNow.ItemID
  )
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (Club.ID)= [Forms]![ClubOrderPrintDialog]![ID]
    )
  );
