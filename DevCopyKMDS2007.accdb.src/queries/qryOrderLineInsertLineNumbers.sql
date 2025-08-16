PARAMETERS parOrdheadID Long,
parLineID Short;
UPDATE
  qryOrderLineByLineIDDescending
SET
  qryOrderLineByLineIDDescending.LineID = [LineID] + 1
WHERE
  (
    (
      (
        qryOrderLineByLineIDDescending.LineID
      )>= [parLineID]
    )
    AND (
      (
        qryOrderLineByLineIDDescending.OrdheadID
      )= [parOrdheadID]
    )
  );
