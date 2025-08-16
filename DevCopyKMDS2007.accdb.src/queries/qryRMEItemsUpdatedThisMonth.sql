SELECT
  InvoLine.Code,
  Sum(
    IIf([Transtype] = "S", [Quantity], 0)
  ) AS [Counter],
  Sum(
    IIf([Transtype] = "C", [Quantity], 0)
  ) AS Clubs,
  Sum(
    IIf([Transtype] = "P", [Quantity], 0)
  ) AS Purchase,
  Sum(
    IIf([Transtype] = "A", [Quantity], 0)
  ) AS OnhandChange,
  Sum(
    IIf([Transtype] = "X", [Quantity], 0)
  ) AS Transfer,
  [Counter] + [Clubs] - [Purchase] - [Transfer] - [OnhandChange] AS Delta
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.SetupDate)> Month(
        Now()
      )& "/01/" & Year(
        Now()
      )
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  )
GROUP BY
  InvoLine.Code;
