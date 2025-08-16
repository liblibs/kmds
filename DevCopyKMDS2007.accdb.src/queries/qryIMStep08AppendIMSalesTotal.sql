INSERT INTO ItemMovement (
  Code, Description, CurrMM, YTDQty,
  Jan, Feb, Mar, Apr, May, Jun, Jul, Aug,
  Sep, Oct, Nov, [Dec]
)
SELECT
  ItemMovement.Code,
  "Sales - Total" AS Descr,
  Sum(ItemMovement.CurrMM) AS SumOfCurrMM,
  Sum(ItemMovement.YTDQty) AS SumOfYTDQty,
  Sum(ItemMovement.Jan) AS SumOfJan,
  Sum(ItemMovement.Feb) AS SumOfFeb,
  Sum(ItemMovement.Mar) AS SumOfMar,
  Sum(ItemMovement.Apr) AS SumOfApr,
  Sum(ItemMovement.May) AS SumOfMay,
  Sum(ItemMovement.Jun) AS SumOfJun,
  Sum(ItemMovement.Jul) AS SumOfJul,
  Sum(ItemMovement.Aug) AS SumOfAug,
  Sum(ItemMovement.Sep) AS SumOfSep,
  Sum(ItemMovement.Oct) AS SumOfOct,
  Sum(ItemMovement.Nov) AS SumOfNov,
  Sum(ItemMovement.Dec) AS SumOfDec
FROM
  ItemMovement
WHERE
  (
    (
      (ItemMovement.Description) Like " Sales*"
    )
  )
GROUP BY
  ItemMovement.Code;
