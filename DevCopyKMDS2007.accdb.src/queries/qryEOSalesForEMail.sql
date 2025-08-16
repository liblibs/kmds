SELECT
  Agents.eMail,
  Sysfile.AgentNumber,
  Sysfile.Name,
  qryEOSalesByDepartmentNoOS.GroupLvl,
  qryEOSalesByDepartmentNoOS.Group,
  IIf(
    [qryEOSalesByDepartmentNoOS].[Transtype] = 'C'
    Or [qryEOSalesByDepartmentNoOS].[Transtype] = 'D',
    [qryEOSalesByDepartmentNoOS].[GroupDesc] & " Bars",
    [qryEOSalesByDepartmentNoOS].[GroupDesc]
  ) AS GroupDescription,
  Sum(
    qryEOSalesByDepartmentNoOS.TotalQ
  ) AS SumOfTotalQ,
  Sum(
    qryEOSalesByDepartmentNoOS.TotalP
  ) AS TotalP,
  Sum(qryEOSalesByDepartmentNoOS.Q1) AS SumOfQ1,
  Sum(qryEOSalesByDepartmentNoOS.D1) AS SumOfD1,
  Sum(qryEOSalesByDepartmentNoOS.Q2) AS SumOfQ2,
  Sum(qryEOSalesByDepartmentNoOS.D2) AS SumOfD2,
  Sum(qryEOSalesByDepartmentNoOS.Q3) AS SumOfQ3,
  Sum(qryEOSalesByDepartmentNoOS.D3) AS SumOfD3,
  Sum(qryEOSalesByDepartmentNoOS.Q4) AS SumOfQ4,
  Sum(qryEOSalesByDepartmentNoOS.D4) AS SumOfD4,
  Sum(
    qryEOSalesByDepartmentNoOS.TotalC
  ) AS SumOfTotalC,
  Sum(
    qryEOSalesByDepartmentNoOS.TotalGP
  ) AS SumOfTotalGP
FROM
  qryEOSalesByDepartmentNoOS,
  Agents
  INNER JOIN Sysfile ON Agents.AgentNumber = Sysfile.AgentNumber
GROUP BY
  Agents.eMail,
  Sysfile.AgentNumber,
  Sysfile.Name,
  qryEOSalesByDepartmentNoOS.GroupLvl,
  qryEOSalesByDepartmentNoOS.Group,
  IIf(
    [qryEOSalesByDepartmentNoOS].[Transtype] = 'C'
    Or [qryEOSalesByDepartmentNoOS].[Transtype] = 'D',
    [qryEOSalesByDepartmentNoOS].[GroupDesc] & " Bars",
    [qryEOSalesByDepartmentNoOS].[GroupDesc]
  )
HAVING
  (
    (
      (
        IIf(
          [qryEOSalesByDepartmentNoOS].[Transtype] = 'C'
          Or [qryEOSalesByDepartmentNoOS].[Transtype] = 'D',
          [qryEOSalesByDepartmentNoOS].[GroupDesc] & " Bars",
          [qryEOSalesByDepartmentNoOS].[GroupDesc]
        )
      )<> "Drops and Loans"
    )
  );
