Operation =3
Name ="UPCLibrary"
Option =0
Where ="(((UPCLibrary.UPCCode) Is Null) And ((Len(UPCItem.ID))>7) And ((Item.SalesDepart"
    "mentID)>1))"
Begin InputTables
    Name ="UPCItem"
    Name ="Item"
    Name ="UPCLibrary"
End
Begin OutputColumns
    Name ="UPCCode"
    Expression ="UPCItem.ID"
    Name ="Description"
    Expression ="Item.Description"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="UPCLibrary"
    Expression ="UPCItem.ID = UPCLibrary.UPCCode"
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
    Top =0
    ColumnsShown =651
    Begin
        Left =402
        Top =8
        Right =498
        Bottom =122
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =592
        Top =10
        Right =688
        Bottom =124
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="UPCLibrary"
        Name =""
    End
End
