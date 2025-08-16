dbMemo "SQL" ="SELECT v.SalesDptDescr, v.CategoryDescr AS Category, i.ID, v.Description AS Desc"
    "ription, s.Description AS [Size], i.RestockAmount, [RestockAmount]*v.Price AS Do"
    "llars, v.OnShelf, i.Location, UCASE(d.[Description]) AS SalesDepartment\015\012F"
    "ROM ((vwItemsWithCurrentPricing AS v INNER JOIN Item AS i ON v.ID=i.ID) INNER JO"
    "IN [Size] AS s ON i.SizeID=s.ID) INNER JOIN SalesDepartment AS d ON v.SalesDepar"
    "tmentID=d.ID\015\012WHERE i.ID Is Not Null And i.RestockAmount<>0\015\012ORDER B"
    "Y v.CategoryDescr, s.Description, i.ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="v.SalesDptDescr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="i.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="i.RestockAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="i.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment"
        dbLong "AggregateType" ="-1"
    End
End
