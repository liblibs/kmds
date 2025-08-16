SELECT
  [olcc order].[ITEM ],
  [olcc order].DESCRIPTION,
  [olcc order].[ORDER ],
  [olcc order].ID,
  Item.ID,
  OLCCitem.Code,
  OLCCitem.Description
FROM
  (
    [olcc order]
    LEFT JOIN Item ON [olcc order].[ITEM ] = Item.ID
  )
  LEFT JOIN OLCCitem ON [olcc order].[ITEM ] = OLCCitem.Code
ORDER BY
  Item.ID;
