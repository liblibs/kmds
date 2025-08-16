Operation =1
Option =0
Where ="(((SalesDepartment.ID)>2))"
Begin InputTables
    Name ="Item"
    Name ="SalesDepartment"
    Name ="qryUPCMRU"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Expression ="Item.ID"
    Alias ="UPC"
    Expression ="qryUPCMRU.ID"
    Expression ="Item.Description"
    Expression ="Item.Price"
    Expression ="Item.Cost"
    Expression ="Item.Deposit"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryUPCMRU"
    Expression ="Item.ID = qryUPCMRU.ItemID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Item.Description"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SalesDepartment.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Deposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryUPCMRU.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-290
    Top =126
    Right =890
    Bottom =932
    Left =-1
    Top =-1
    Right =1148
    Bottom =522
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =244
        Bottom =389
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =329
        Top =9
        Right =473
        Bottom =367
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =521
        Top =12
        Right =665
        Bottom =156
        Top =0
        Name ="qryUPCMRU"
        Name =""
    End
End
