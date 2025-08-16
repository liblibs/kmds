INSERT INTO Audit (
  ItemID, OnShelf, PhysicalCount, VarianceAmount,
  Price, Cost, CountFront
)
SELECT
  Item.ID,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.OnShelf AS PhysicalCount,
  0 AS VarianceAmount,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.Cost,
  vwItemsWithCurrentPricing.OnShelf AS Front
FROM
  (
    (
      Audit
      RIGHT JOIN Item ON Audit.ItemID = Item.ID
    )
    INNER JOIN vwItemsWithCurrentPricing ON Item.ID = vwItemsWithCurrentPricing.ID
  )
  INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
WHERE
  (
    (
      (Audit.ItemID) Is Null
    )
    And (
      (SubCategory.Description) Between Forms!Audit!cmbSubCategoryID
      And Forms!Audit!cmbSubCategoryIDEnd
    )
  );
