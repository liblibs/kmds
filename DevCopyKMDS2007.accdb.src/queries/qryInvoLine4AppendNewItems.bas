dbMemo "SQL" ="INSERT INTO Item ( ID, SalesDepartmentID, VendorID, Description, Price, SetupDat"
    "e, ChangeDate, SizeID, UnitsPerCase, UnitsPerOrder, OrderTo )\015\012SELECT Invo"
    "Line4.Code AS ID, 1 AS SalesDepartmentID, 1 AS VendorID, \"= Setup from EOD\" AS"
    " Description, InvoLine4.Price, Now() AS SetupDate, Now() AS ChangeDate, IIf(IsNu"
    "ll(Size.id),0,size.id) AS SizeID, IIf(IsNull(Size.CaseQty),0,size.CaseQty) AS Un"
    "itsPerCase, 1 AS UnitsPerOrder, 1 AS OrderTo\015\012FROM (InvoLine4 LEFT JOIN It"
    "em ON InvoLine4.Code=Item.ID) LEFT JOIN [Size] ON InvoLine4.Drawer=Size.ID\015\012"
    "WHERE (((InvoLine4.TransType)=\"S\") AND ((Item.ID) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderTo"
        dbLong "AggregateType" ="-1"
    End
End
