SELECT
  IIf([SalesDepartmentID] = 1,-1, 0) AS LiquorItem,
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  Club.ID AS ClubID,
  InvoLine.InvoHeadID AS InvoiceNo,
  InvoLine.SetupDate AS OrdDate,
  InvoLine.SetupDate AS InvDate,
  InvoLine.Quantity,
  InvoLine.Cost AS ExtCost,
  1 AS LiquorExtPrice,
  1 AS LiquorDiscount,
  [LiquorExtPrice] - [LiquorDiscount] AS LiquorNetPrice,
  1 AS NonLiquorExtPrice,
  1 AS NonLiquorDiscount,
  [NonLiquorExtPrice] - [NonLiquorDiscount] AS NonLiquorNetPrice,
  Club.NonLiquor,
  Club.Name AS ClubName,
  Club.Contact AS ClubContact,
  Club.Address1 AS ClubAddress1,
  Club.Address2 AS ClubAddress2,
  [City] & ", " & [ST] & "  " & [Zip] AS ClubCityST,
  Club.Phone1 AS ClubPhone1,
  InvoLine.ID AS LineID,
  InvoLine.Code AS ItemID,
  Item.Description,
  Size.Description AS SizeDesc,
  InvoLine.Quantity AS QuantityShipped,
  InvoLine.Price,
  InvoLine.Cost,
  Club.ClubNotes,
  "" AS SalePrice,
  "" AS Buydown,
  InvoLine.Price AS NetPrice,
  IIf(
    (
      IsNull(size.Oz)
      Or size.oz = 0
    ),
    1,
    size.oz
  ) AS Ounces,
  CCur(
    IIf(
      [involine].[transtype] = "D",
      0,
      (
        [InvoLine].[price] / [InvoLine].[Quantity]
      )/ [Ounces]
    )
  ) AS PricePerOZ,
  Club.Phone2 AS ClubFax,
  Item.SalesDepartmentID,
  SubCategory.Description AS CatDesc,
  Item.Location,
  InvoLine.SalesGroup AS InvoiceNumber
FROM
  Club
  INNER JOIN (
    InvoLine
    INNER JOIN (
      (
        Item
        LEFT JOIN [Size] ON Item.SizeID = Size.ID
      )
      LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
    ) ON InvoLine.Code = Item.ID
  ) ON Club.ID = InvoLine.SalesDept
WHERE
  (
    (
      (InvoLine.CashSysID)= Forms!Clubs!ClubsPreviousOrderSub.Form!CashSysID
    )
    And (
      (InvoLine.InvoHeadID)= Forms!Clubs!ClubsPreviousOrderSub.Form!InvoHeadID
    )
    And (
      (InvoLine.SetupDate)= Forms!Clubs!ClubsPreviousOrderSub.Form!PostDate
    )
    And (
      (InvoLine.Quantity)<> 0
      And (InvoLine.Quantity)<> 0
    )
  );
