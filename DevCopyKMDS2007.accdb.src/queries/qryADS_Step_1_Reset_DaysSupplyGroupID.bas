Operation =4
Option =0
Where ="(((Item.DaysSupplyGroupID)<>0 Or (Item.DaysSupplyGroupID) Is Null)) OR (((Item.A"
    "verageDailySales)<>0 Or (Item.AverageDailySales) Is Null))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.DaysSupplyGroupID"
    Expression ="0"
    Name ="Item.AverageDailySales"
    Expression ="0"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =170
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
