Operation =1
Option =0
Where ="(((PriceGroup.ID)<>0))"
Begin InputTables
    Name ="Item"
    Name ="PriceGroup"
End
Begin OutputColumns
    Expression ="PriceGroup.Description"
    Expression ="Item.Description"
    Expression ="Item.Price"
    Expression ="Item.Cost"
    Expression ="Item.Deposit"
    Expression ="Item.UnitsPerCase"
    Expression ="Item.QtyPack"
    Expression ="Item.QtyPackPrice"
    Expression ="PriceGroup.ID"
    Expression ="Item.ID"
    Alias ="GP"
    Expression ="IIf([PRICE]>0 And [COST]>0,([PRICE]-[COST])/[PRICE],0)"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="PriceGroup"
    Expression ="Item.PriceGroupID = PriceGroup.ID"
    Flag =3
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
        dbText "Name" ="PriceGroup.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
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
        dbText "Name" ="Item.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.QtyPackPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriceGroup.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GP"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =327
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =304
        Top =36
        Right =615
        Bottom =376
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =26
        Top =100
        Right =170
        Bottom =244
        Top =0
        Name ="PriceGroup"
        Name =""
    End
End
