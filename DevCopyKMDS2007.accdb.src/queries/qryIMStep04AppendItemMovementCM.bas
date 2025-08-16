Operation =3
Name ="ItemMovementCurrentMonth"
Option =0
Where ="(((InvoLine.SetupDate)>=[parDate]))"
Begin InputTables
    Name ="InvoLine"
    Name ="TranTypes"
End
Begin OutputColumns
    Name ="Code"
    Expression ="InvoLine.Code"
    Name ="Description"
    Expression ="TranTypes.Description"
    Alias ="CurMQty"
    Name ="CurMQty"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="CurMDollars"
    Name ="CurMDollars"
    Expression ="Sum(InvoLine.Price)"
End
Begin Parameters
    Name ="parDate"
    Flag =8
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType = TranTypes.Code"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="TranTypes.Description"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbBoolean "UseTransaction" ="-1"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="CurMQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurMDollars"
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
    Bottom =211
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =180
        Top =7
        Right =310
        Bottom =196
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =51
        Top =51
        Right =147
        Bottom =135
        Top =0
        Name ="TranTypes"
        Name =""
    End
End
