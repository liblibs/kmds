SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  InvoLine.ID,
  InvoLine.Code,
  IIf(
    [TransType] = "T", Tender.Description,
    Item.Description
  ) AS CodeDesc,
  IIf(
    [TransType] = "S"
    Or [TransType] = "C",
    [Quantity], 0
  ) AS Qty,
  CCur(
    InvoLine.Price / IIf([quantity] = 0, 1, [quantity])
  ) AS UPrice,
  IIf(
    [TransType] = "T", 0 - InvoLine.Price,
    InvoLine.Price
  ) AS Extend,
  InvoLine.SalesDept,
  InvoLine.InputCode,
  InvoLine.SalesGroup,
  InvoLine.Drawer,
  InvoLine.TransType,
  InvoLine.UpdateOnhand,
  Cashier.Name AS CashierName,
  IIf([TransType] = "C", Club.Name, "") AS ClubName,
  TranTypes.Description AS [Transaction]
FROM
  (
    (
      (
        (
          (
            InvoLine
            LEFT JOIN Item ON InvoLine.Code = Item.ID
          )
          LEFT JOIN Cashier ON InvoLine.CashierID = Cashier.ID
        )
        LEFT JOIN Club ON InvoLine.SalesDept = Club.ID
      )
      LEFT JOIN Tender ON InvoLine.Code = Tender.ID
    )
    LEFT JOIN Agents ON InvoLine.SalesDept = Agents.AgentNumber
  )
  LEFT JOIN TranTypes ON InvoLine.TransType = TranTypes.Code
WHERE
  (
    (
      (InvoLine.InvoHeadID)= Forms!ItemTransactionsByMonthPO!InvoHeadID
    )
    And (
      (InvoLine.CashSysID)= Forms!ItemTransactionsByMonthPO!CashSysID
    )
  );
