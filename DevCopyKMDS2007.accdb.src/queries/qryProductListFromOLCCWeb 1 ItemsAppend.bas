dbMemo "SQL" ="INSERT INTO item ( ID, SizeID, SalesDepartmentID, VendorID, Description, Price, "
    "UnitsPerCase, UnitsPerOrder )\015\012SELECT qryProductListFromOLCCWebsite.Code, "
    "qryProductListFromOLCCWebsite.Size, 1 AS sd, 1 AS VendorID, qryProductListFromOL"
    "CCWebsite.Description, qryProductListFromOLCCWebsite.Price, qryProductListFromOL"
    "CCWebsite.CaseQ, qryProductListFromOLCCWebsite.CaseQ\015\012FROM Item RIGHT JOIN"
    " qryProductListFromOLCCWebsite ON Item.ID=qryProductListFromOLCCWebsite.Code\015"
    "\012WHERE (((Item.ID) Is Null));\015\012"
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
        dbText "Name" ="Item.Price"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
End
