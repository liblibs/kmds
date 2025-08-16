Operation =4
Option =0
Where ="(((PalmFrom.Physical) Is Not Null))"
Begin InputTables
    Name ="PalmFrom"
    Name ="Audit"
End
Begin OutputColumns
    Name ="Audit.PhysicalCount"
    Expression ="PalmFrom.Physical"
    Name ="Audit.VarianceAmount"
    Expression ="PalmFrom.Physical-Audit.OnShelf"
    Name ="Audit.CountFront"
    Expression ="PalmFrom.Physical"
    Name ="Audit.CountBack"
    Expression ="0"
    Name ="Audit.CountDisplay"
    Expression ="0"
    Name ="Audit.CountOverflow"
    Expression ="0"
End
Begin Joins
    LeftTable ="PalmFrom"
    RightTable ="Audit"
    Expression ="PalmFrom.ItemID = Audit.ItemID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
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
    Bottom =192
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =150
        Top =0
        Name ="PalmFrom"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =348
        Bottom =180
        Top =0
        Name ="Audit"
        Name =""
    End
End
