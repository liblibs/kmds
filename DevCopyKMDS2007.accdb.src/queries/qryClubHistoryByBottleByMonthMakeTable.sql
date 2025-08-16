SELECT
  Club.ID AS ClubID,
  Club.Name AS ClubName,
  IIf(
    IsNull([subcategory].[Description]),
    " Not Assigned",
    [subcategory].[Description]
  ) AS CatDesc,
  InvoLine.Code,
  IIf(
    IsNull(Item.Description),
    "Deleted",
    Item.Description
  ) AS Description,
  Size.Description AS SizeD,
  Format(InvoLine.SetupDate, "mm") AS [Year-Month],
  Sum(InvoLine.Quantity) AS Qty,
  Sum(InvoLine.Price) AS Dollars INTO ClubHistoryByBottleByMonth
FROM
  (
    Club
    INNER JOIN (
      qryClubInvoiceHistoryByInvoice
      INNER JOIN (
        (
          InvoLine
          LEFT JOIN Item ON InvoLine.Code = Item.ID
        )
        LEFT JOIN [Size] ON Item.SizeID = Size.ID
      ) ON (
        qryClubInvoiceHistoryByInvoice.InvoHeadID = InvoLine.InvoHeadID
      )
      AND (
        qryClubInvoiceHistoryByInvoice.CashSysID = InvoLine.CashSysID
      )
      AND (
        qryClubInvoiceHistoryByInvoice.SetupDate = InvoLine.SetupDate
      )
    ) ON Club.ID = qryClubInvoiceHistoryByInvoice.ClubID
  )
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  )
GROUP BY
  Club.ID,
  Club.Name,
  IIf(
    IsNull([subcategory].[Description]),
    " Not Assigned",
    [subcategory].[Description]
  ),
  InvoLine.Code,
  IIf(
    IsNull(Item.Description),
    "Deleted",
    Item.Description
  ),
  Size.Description,
  Format(InvoLine.SetupDate, "mm");
