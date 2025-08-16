SELECT
  Audit.ItemID AS Code,
  Item.Description,
  IIf(
    IsNull([FirstOfID]),
    Audit.ItemID,
    [FirstOfID]
  ) AS UPCCode,
  Audit.OnShelf INTO PalmTo
FROM
  (
    Audit
    INNER JOIN Item ON Audit.ItemID = Item.ID
  )
  LEFT JOIN qryAuditFileToPalm_Step_2a ON Item.ID = qryAuditFileToPalm_Step_2a.ItemID
ORDER BY
  IIf(
    IsNull([FirstOfID]),
    Audit.ItemID,
    [FirstOfID]
  );
