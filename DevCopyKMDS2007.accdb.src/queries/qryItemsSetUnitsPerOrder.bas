Operation =4
Option =0
Where ="(((Item.UnitsPerOrder) Is Null Or (Item.UnitsPerOrder)<1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.UnitsPerOrder"
    Expression ="1"
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
    Bottom =77
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =517
        Top =6
        Right =613
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
