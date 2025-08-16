Operation =3
Name ="UPCItem"
Option =0
Where ="((([Query13].[NewItemID])>0))"
Begin InputTables
    Name ="Query13"
End
Begin OutputColumns
    Alias ="ne"
    Name ="ID"
    Expression ="checkdigit11(Right([newUPC2],11))"
    Alias ="Expr1"
    Name ="ItemID"
    Expression ="Query13.NewItemID"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Query11.NewUPC1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query11.lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query11.le"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ne"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query11.NewItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query13.NewItemID"
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
    Bottom =178
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Query13"
        Name =""
    End
End
