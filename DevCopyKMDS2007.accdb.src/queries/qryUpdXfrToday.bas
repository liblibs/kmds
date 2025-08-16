Operation =4
Option =0
Where ="(((Item.ID)=[parItemID]))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.TransferredToday"
    Expression ="[TransferredToday]+[parSold]"
End
Begin Parameters
    Name ="parItemID"
    Flag =4
    Name ="parSold"
    Flag =3
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
        Right =134
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
