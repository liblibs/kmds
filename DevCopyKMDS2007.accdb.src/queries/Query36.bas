Operation =3
Name ="ITEM"
Option =0
Begin InputTables
    Name ="WilsonvilleOrderNOF"
    Name ="dbo_OLCCitem"
End
Begin OutputColumns
    Name ="ID"
    Expression ="dbo_OLCCitem.Code"
    Name ="Description"
    Expression ="dbo_OLCCitem.Description"
    Name ="Price"
    Expression ="dbo_OLCCitem.Price"
    Name ="UnitsPerCase"
    Expression ="dbo_OLCCitem.UnitsPerCase"
    Name ="StockNumber"
    Expression ="dbo_OLCCitem.StockStatus"
    Alias ="SD"
    Name ="SizeID"
    Expression ="Right([CODE],1)"
    Alias ="SLD"
    Name ="SalesDepartmentID"
    Expression ="1"
    Alias ="VD"
    Name ="VendorID"
    Expression ="1"
    Name ="Cost"
    Expression ="dbo_OLCCitem.Price"
End
Begin Joins
    LeftTable ="WilsonvilleOrderNOF"
    RightTable ="dbo_OLCCitem"
    Expression ="WilsonvilleOrderNOF.ItemID = dbo_OLCCitem.Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCitem.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.StockStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WilsonvilleOrder.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WilsonvilleOrder.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SLD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VD"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-4
    Top =-23
    Right =1178
    Bottom =877
    Left =-1
    Top =-1
    Right =1159
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =106
        Top =49
        Right =250
        Bottom =193
        Top =0
        Name ="WilsonvilleOrderNOF"
        Name =""
    End
    Begin
        Left =435
        Top =41
        Right =579
        Bottom =355
        Top =0
        Name ="dbo_OLCCitem"
        Name =""
    End
End
