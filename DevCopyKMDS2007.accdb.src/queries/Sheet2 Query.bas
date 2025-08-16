dbMemo "SQL" ="INSERT INTO Item ( ID, Description, Price, SizeId, SalesDepartmentID, VendorID, "
    "Cost )\015\012SELECT Sheet2.[Code], Sheet2.[Description], Sheet2.[######] AS Pri"
    "ce, Right([code],1) AS SizeId, 1 AS SalesDepartmentID, 1 AS VendorID, Sheet2.[##"
    "####] AS Cost\015\012FROM Sheet2 LEFT JOIN Item ON Sheet2.Code = Item.ID\015\012"
    "WHERE (((Item.ID) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="[Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Description]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet2.[Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet2.[Description]"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet2.[######]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
End
