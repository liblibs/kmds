Operation =3
Name ="ItemMovement"
Option =0
Where ="(((ItemMovement.Code) Is Null))"
Begin InputTables
    Name ="qryALLItemsAndALLTransTypes"
    Name ="ItemMovement"
End
Begin OutputColumns
    Name ="Code"
    Expression ="qryALLItemsAndALLTransTypes.ID"
    Name ="Description"
    Expression ="qryALLItemsAndALLTransTypes.Description"
End
Begin Joins
    LeftTable ="qryALLItemsAndALLTransTypes"
    RightTable ="ItemMovement"
    Expression ="qryALLItemsAndALLTransTypes.ID = ItemMovement.Code"
    Flag =2
    LeftTable ="qryALLItemsAndALLTransTypes"
    RightTable ="ItemMovement"
    Expression ="qryALLItemsAndALLTransTypes.Description = ItemMovement.Description"
    Flag =2
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
    Bottom =110
    Left =0
    Top =13
    ColumnsShown =651
    Begin
        Left =27
        Top =7
        Right =312
        Bottom =106
        Top =0
        Name ="qryALLItemsAndALLTransTypes"
        Name =""
    End
    Begin
        Left =375
        Top =7
        Right =540
        Bottom =121
        Top =0
        Name ="ItemMovement"
        Name =""
    End
End
