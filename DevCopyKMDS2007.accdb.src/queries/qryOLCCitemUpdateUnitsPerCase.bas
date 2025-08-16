Operation =4
Option =0
Where ="(((Item.UnitsPerCase)<>[dbo_OLCCitem].[UnitsPerCase]))"
Begin InputTables
    Name ="dbo_OLCCitem"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.UnitsPerCase"
    Expression ="[dbo_OLCCitem].[UnitsPerCase]"
End
Begin Joins
    LeftTable ="dbo_OLCCitem"
    RightTable ="Item"
    Expression ="dbo_OLCCitem.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCitem.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =88
    Right =1271
    Bottom =829
    Left =-1
    Top =-1
    Right =1200
    Bottom =457
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =77
        Top =31
        Right =221
        Bottom =434
        Top =0
        Name ="dbo_OLCCitem"
        Name =""
    End
    Begin
        Left =313
        Top =38
        Right =457
        Bottom =411
        Top =0
        Name ="Item"
        Name =""
    End
End
