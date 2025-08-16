UPDATE
  Item
  INNER JOIN qryProductListFromOLCCWebsite ON Item.ID = qryProductListFromOLCCWebsite.Code
SET
  Item.Description = qryProductListFromOLCCWebsite.description,
  Item.ChangeDate = Now()
WHERE
  (
    (
      (
        qryProductListFromOLCCWebsite.Description
      )<> item.description
    )
  );
