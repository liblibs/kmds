SELECT
  Club.ID AS ClubID,
  Club.Name AS ClubName,
  SubCategory.Description,
  Item.Description,
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS TotalQty,
  Sum(InvoLine.Cost) AS TotalCost,
  Sum(
    IIf(
      [TransType] = 'C', [involine].[Price],
      0
    )
  ) AS Retail,
  Sum(
    IIf(
      [TransType] = 'D', [involine].[Price],
      0
    )
  ) AS Discount,
  Sum(InvoLine.Price) AS TotalPrice
FROM
  (
    (
      InvoLine
      INNER JOIN Club ON InvoLine.SalesDept = Club.ID
    )
    INNER JOIN Item ON InvoLine.Code = Item.ID
  )
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (InvoLine.TransType)= "C"
      Or (InvoLine.TransType)= "D"
    )
  )
GROUP BY
  Club.ID,
  Club.Name,
  SubCategory.Description,
  Item.Description,
  InvoLine.Code
HAVING
  (
    (
      (Club.ID)= [Forms]![ClubOrderWebImport]![cmbLicensee]
    )
  )
ORDER BY
  SubCategory.Description,
  Item.Description;
