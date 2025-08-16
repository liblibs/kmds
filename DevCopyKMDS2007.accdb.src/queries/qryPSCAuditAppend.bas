Operation =3
Name ="Audit"
Option =0
Begin InputTables
    Name ="qryPSCAuditInventory"
End
Begin OutputColumns
    Alias ="ItemID"
    Name ="ItemID"
    Expression ="qryPSCAuditInventory.ID"
    Name ="OnShelf"
    Expression ="qryPSCAuditInventory.OnShelf"
    Alias ="qs"
    Name ="PhysicalCount"
    Expression ="IIf([QuantityShipped]<0,[OnShelf],[QuantityShipped])"
    Alias ="VarianceAmount"
    Name ="VarianceAmount"
    Expression ="(IIf([QuantityShipped]<0,[OnShelf],[QuantityShipped]))-[onshelf]"
    Name ="Price"
    Expression ="qryPSCAuditInventory.CurPrice"
    Name ="Cost"
    Expression ="qryPSCAuditInventory.Cost"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VarianceAmount"
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
    Bottom =161
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =165
        Top =0
        Name ="qryPSCAuditInventory"
        Name =""
    End
End
