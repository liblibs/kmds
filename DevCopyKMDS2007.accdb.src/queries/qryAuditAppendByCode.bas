Operation =3
Name ="Audit"
Option =0
Where ="(((Item.ID) Between Forms!Audit!cmbSubCategoryID And Forms!Audit!cmbSubCategoryI"
    "DEnd) And ((Audit.ItemID) Is Null))"
Begin InputTables
    Name ="Audit"
    Name ="Item"
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="Item.ID"
    Name ="OnShelf"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Alias ="PhysicalCount"
    Name ="PhysicalCount"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Alias ="VarianceAmount"
    Name ="VarianceAmount"
    Expression ="0"
    Name ="Price"
    Expression ="vwItemsWithCurrentPricing.Price"
    Name ="Cost"
    Expression ="vwItemsWithCurrentPricing.Cost"
    Alias ="Front"
    Name ="CountFront"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
End
Begin Joins
    LeftTable ="Audit"
    RightTable ="Item"
    Expression ="Audit.ItemID = Item.ID"
    Flag =3
    LeftTable ="Item"
    RightTable ="vwItemsWithCurrentPricing"
    Expression ="Item.ID = vwItemsWithCurrentPricing.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="PhysicalCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VarianceAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Front"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =18
        Top =5
        Right =114
        Bottom =119
        Top =0
        Name ="Audit"
        Name =""
    End
    Begin
        Left =175
        Top =6
        Right =271
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =330
        Top =10
        Right =600
        Bottom =124
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
