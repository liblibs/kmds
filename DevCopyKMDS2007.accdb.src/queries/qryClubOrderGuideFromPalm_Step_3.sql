SELECT
  PalmOrderGuideFrom.ItemID,
  Min(PalmOrderGuideFrom.Physical) AS MinOfPhysical,
  IIf([SalesDepartmentID] = 1, 255, 0) AS Liquor
FROM
  PalmOrderGuideFrom
  INNER JOIN Item ON PalmOrderGuideFrom.ItemID = Item.ID
GROUP BY
  PalmOrderGuideFrom.ItemID,
  IIf([SalesDepartmentID] = 1, 255, 0)
HAVING
  (
    (
      (
        Min(PalmOrderGuideFrom.Physical)
      ) Is Not Null
      And (
        Min(PalmOrderGuideFrom.Physical)
      )> 0
    )
  )
ORDER BY
  PalmOrderGuideFrom.ItemID;
