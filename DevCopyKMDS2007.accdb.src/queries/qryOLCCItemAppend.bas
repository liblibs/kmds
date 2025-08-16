Operation =3
Name ="dbo_OLCCitem"
Option =0
Where ="(((Item_ICExpansion.Code) Is Not Null))"
Begin InputTables
    Name ="Item_ICExpansion"
End
Begin OutputColumns
    Name ="Code"
    Expression ="Item_ICExpansion.Code"
    Name ="Description"
    Expression ="Item_ICExpansion.Description"
    Name ="Age"
    Expression ="Item_ICExpansion.Age"
    Name ="Proof"
    Expression ="Item_ICExpansion.Proof"
    Name ="Price"
    Expression ="Item_ICExpansion.Price"
    Name ="UnitsPerCase"
    Expression ="Item_ICExpansion.UnitsPerCase"
    Name ="StockStatus"
    Expression ="Item_ICExpansion.StockStatus"
    Name ="OldItemID"
    Expression ="Item_ICExpansion.OldItemID"
    Name ="OldItemAlpha"
    Expression ="Item_ICExpansion.OldItemAlpha"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="OLCCitem.StockStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Age"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.Age"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Proof"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.Proof"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.StockStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.OldItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_ICExpansion.OldItemAlpha"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-111
    Top =36
    Right =1114
    Bottom =713
    Left =-1
    Top =-1
    Right =1193
    Bottom =303
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =382
        Top =10
        Right =553
        Bottom =266
        Top =0
        Name ="Item_ICExpansion"
        Name =""
    End
End
