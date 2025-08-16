Operation =4
Option =0
Where ="(((InvoLine5.Code)=[DepositItemID]) AND ((Item.Price)>0))"
Begin InputTables
    Name ="Sysfile"
    Name ="qryBottleDepositQtyAdjAtEOD"
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Name ="InvoLine5.Quantity"
    Expression ="[Involine5].[Price]/[item].[price]"
    Name ="InvoLine5.Cost"
    Expression ="[InvoLine5].[PRICE]"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code=Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Sysfile.DepositItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Cost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =213
    Top =259
    Right =1779
    Bottom =1672
    Left =-1
    Top =-1
    Right =1528
    Bottom =832
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =623
        Top =19
        Right =803
        Bottom =1052
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =851
        Top =30
        Right =1031
        Bottom =210
        Top =0
        Name ="qryBottleDepositQtyAdjAtEOD"
        Name =""
    End
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =484
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =329
        Top =34
        Right =558
        Bottom =424
        Top =0
        Name ="Item"
        Name =""
    End
End
