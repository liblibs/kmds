dbMemo "SQL" ="INSERT INTO item ( ID, Description, SalesDepartmentID, Onhand, Cost, Price )\015"
    "\012SELECT MeadowlandInventoryValuex.NewItem, MeadowlandInventoryValuex.descript"
    "ion, MeadowlandInventoryValuex.SalesDpt AS SalesDepartmentID, MeadowlandInventor"
    "yValuex.Onhand, MeadowlandInventoryValuex.Cost, MeadowlandInventoryValuex.Price\015"
    "\012FROM MeadowlandInventoryValuex LEFT JOIN SalesDepartment ON MeadowlandInvent"
    "oryValuex.SalesDpt=SalesDepartment.ID\015\012WHERE (((MeadowlandInventoryValuex."
    "NewItem)>0));\015\012"
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
        dbText "Name" ="MeadowlandInventoryValuex.ItemNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.NewItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.SalesDpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeadowlandInventoryValuex.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ds"
        dbLong "AggregateType" ="-1"
    End
End
