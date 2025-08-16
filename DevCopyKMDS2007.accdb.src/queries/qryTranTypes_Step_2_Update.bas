Operation =4
Option =0
Begin InputTables
    Name ="TranTypes"
    Name ="TranTypesLocal"
End
Begin OutputColumns
    Name ="TranTypes.Description"
    Expression ="TranTypesLocal.description"
End
Begin Joins
    LeftTable ="TranTypes"
    RightTable ="TranTypesLocal"
    Expression ="TranTypes.Code = TranTypesLocal.Code"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =105
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =388
        Bottom =90
        Top =0
        Name ="TranTypesLocal"
        Name =""
    End
End
