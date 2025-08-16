Operation =4
Option =0
Begin InputTables
    Name ="PricingWorkTable"
End
Begin OutputColumns
    Name ="PricingWorkTable.NewPrice"
    Expression ="Round(([Price]+([Forms]![PricingTool]![txtPercent]*[PricingWorkTable].[Price])*0"
        ".01),2)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="PricingWorkTable.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricingWorkTable.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ex3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewPrice"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewPricex"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricingWorkTable.NewPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PricingWorkTable.NewGP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewGP"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =153
    Right =988
    Bottom =959
    Left =-1
    Top =-1
    Right =868
    Bottom =420
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =246
        Top =0
        Name ="PricingWorkTable"
        Name =""
    End
End
