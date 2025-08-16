Operation =5
Option =0
Where ="(((OrdHead.ID)=[parOrdHeadID]))"
Begin InputTables
    Name ="OrdHead"
End
Begin OutputColumns
    Expression ="OrdHead.ID"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="OrdHead"
        Name =""
    End
End
