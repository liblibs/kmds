Operation =3
Name ="Item"
Option =0
Where ="(((Item.ID) Is Null))"
Begin InputTables
    Name ="GaribaldiSalesHist"
    Name ="Item"
    Name ="dbo_OLCCitem"
End
Begin OutputColumns
    Name ="ID"
    Expression ="GaribaldiSalesHist.ItemId"
    Name ="Description"
    Expression ="dbo_OLCCitem.Description"
    Alias ="SizeID"
    Name ="SizeID"
    Expression ="Right([itemid],1)"
    Alias ="SalesDepartmentID"
    Name ="SalesDepartmentID"
    Expression ="1"
    Alias ="VendorID"
    Name ="VendorID"
    Expression ="1"
    Name ="Price"
    Expression ="dbo_OLCCitem.Price"
    Name ="UnitsPerCase"
    Expression ="dbo_OLCCitem.UnitsPerCase"
End
Begin Joins
    LeftTable ="GaribaldiSalesHist"
    RightTable ="Item"
    Expression ="GaribaldiSalesHist.ItemId = Item.ID"
    Flag =2
    LeftTable ="GaribaldiSalesHist"
    RightTable ="dbo_OLCCitem"
    Expression ="GaribaldiSalesHist.ItemId = dbo_OLCCitem.Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribaldiSalesHist.ItemId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribaldUnitsPerCase.Description"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GaribaldUnitsPerCase.Case"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.Description"
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
    Begin
        dbText "Name" ="dbo_OLCCitem.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-371
    Top =105
    Right =1056
    Bottom =870
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
        Right =192
        Bottom =156
        Top =0
        Name ="GaribaldiSalesHist"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =394
        Bottom =410
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =524
        Top =187
        Right =668
        Bottom =410
        Top =0
        Name ="dbo_OLCCitem"
        Name =""
    End
End
