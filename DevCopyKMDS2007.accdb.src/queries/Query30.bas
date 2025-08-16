Operation =3
Name ="Item"
Option =0
Where ="(((Item.ID) Is Null))"
Begin InputTables
    Name ="Item"
    Name ="GaribalidiInventory"
End
Begin OutputColumns
    Alias ="ID"
    Name ="ID"
    Expression ="GaribalidiInventory.ItemID"
    Name ="Description"
    Expression ="GaribalidiInventory.Description"
    Name ="Price"
    Expression ="GaribalidiInventory.Price"
    Alias ="SizeID"
    Name ="SizeID"
    Expression ="Right([GaribalidiInventory].[itemid],1)"
    Alias ="SalesDepartmentID"
    Name ="SalesDepartmentID"
    Expression ="1"
    Alias ="VendorID"
    Name ="VendorID"
    Expression ="1"
End
Begin Joins
    LeftTable ="GaribalidiInventory"
    RightTable ="Item"
    Expression ="GaribalidiInventory.ItemID = Item.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="GaribalidiInventory.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribalidiInventory.Description"
        dbInteger "ColumnWidth" ="3945"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribalidiInventory.Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribalidiInventory.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeID"
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
End
Begin
    State =0
    Left =38
    Top =86
    Right =1465
    Bottom =851
    Left =-1
    Top =-1
    Right =1395
    Bottom =482
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =262
        Bottom =293
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =376
        Top =16
        Right =520
        Bottom =160
        Top =0
        Name ="GaribalidiInventory"
        Name =""
    End
End
