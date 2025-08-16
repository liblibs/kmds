Operation =1
Option =0
Where ="(((IIf(Forms!Audit!chkOnlyVariance,-1,0))=0)) Or (((Audit.VarianceAmount)<>0) An"
    "d ((IIf(Forms!Audit!chkOnlyVariance,-1,0))=-1))"
Begin InputTables
    Name ="Audit"
    Name ="Item"
    Name ="Size"
End
Begin OutputColumns
    Expression ="Audit.ItemID"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Expression ="Audit.OnShelf"
    Expression ="Audit.CountFront"
    Expression ="Audit.CountBack"
    Expression ="Audit.CountDisplay"
    Expression ="Audit.CountOverflow"
    Expression ="Audit.PhysicalCount"
    Alias ="Var"
    Expression ="[PhysicalCount]-[onshelf]"
    Expression ="Audit.Price"
    Expression ="Audit.Cost"
    Expression ="Audit.SetupDate"
    Expression ="Audit.ChangeDate"
    Alias ="ExtPrice"
    Expression ="([PhysicalCount]-[onshelf])*Audit.Price"
    Alias ="ExtCost"
    Expression ="([PhysicalCount]-[onshelf])*Audit.Cost"
    Expression ="Audit.VarianceAmount"
End
Begin Joins
    LeftTable ="Audit"
    RightTable ="Item"
    Expression ="Audit.ItemID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
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
        dbText "Name" ="SizeDesc"
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
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="Size"
        Name =""
    End
End
