Operation =1
Option =0
Begin InputTables
    Name ="Audit"
End
Begin OutputColumns
    Expression ="Audit.ItemID"
    Expression ="Audit.OnShelf"
    Expression ="Audit.CountFront"
    Expression ="Audit.CountBack"
    Expression ="Audit.PhysicalCount"
    Expression ="Audit.CountDisplay"
    Expression ="Audit.CountOverflow"
    Alias ="Var"
    Expression ="([CountFront]+[CountBack]+[CountDisplay]+[CountOverflow])-[onshelf]"
    Expression ="Audit.Price"
    Expression ="Audit.Cost"
    Expression ="Audit.SetupDate"
    Expression ="Audit.ChangeDate"
    Alias ="ExtPrice"
    Expression ="([Var])*[Price]"
    Alias ="ExtCost"
    Expression ="([Var])*[Cost]"
    Expression ="Audit.VarianceAmount"
    Expression ="Audit.ID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Var"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.OnShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.CountFront"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.CountBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.PhysicalCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.CountDisplay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.CountOverflow"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.VarianceAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Audit.ID"
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
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Audit"
        Name =""
    End
End
