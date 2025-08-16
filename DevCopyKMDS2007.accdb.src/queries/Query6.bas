dbMemo "SQL" ="INSERT INTO Item ( ID, Description, SizeID, SalesDepartmentID, VendorID, Price, "
    "Cost )\015\012SELECT Sheet3.Code AS ID1, Sheet3.descr AS Description, Right([cod"
    "e],1) AS SizeID, 1 AS SalesDepartmentID, 1 AS VendorID, Sheet3.Price, Sheet3.Pri"
    "ce\015\012FROM Item RIGHT JOIN Sheet3 ON Item.ID=Sheet3.Code\015\012WHERE (((Ite"
    "m.ID) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Sheet3.descr"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeID"
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
End
