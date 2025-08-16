PARAMETERS parOrdheadID Long,
parLineID Short,
parItemID IEEEDouble; INSERT INTO OrderLine (LineID, OrdheadID, ItemID)
SELECT
  [parLineID] AS xLineID,
  [parOrdheadID] AS xOrdheadID,
  [parItemID] AS xItemId;
