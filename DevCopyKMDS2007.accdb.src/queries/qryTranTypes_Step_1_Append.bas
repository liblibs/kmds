Operation =3
Name ="TranTypes"
Option =0
Where ="(((TranTypes.Code) Is Null))"
Begin InputTables
    Name ="TranTypes"
    Name ="TranTypesLocal"
End
Begin OutputColumns
    Name ="Code"
    Expression ="TranTypesLocal.Code"
    Name ="Description"
    Expression ="TranTypesLocal.Description"
End
Begin Joins
    LeftTable ="TranTypes"
    RightTable ="TranTypesLocal"
    Expression ="TranTypes.Code = TranTypesLocal.Code"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
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
    ColumnsShown =651
    Begin
        Left =326
        Top =5
        Right =422
        Bottom =104
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =90
        Top =0
        Name ="TranTypesLocal"
        Name =""
    End
End
