Operation =1
Option =0
Where ="(((ItemMovement.Description)=\" Sales - Counter\" Or (ItemMovement.Description)="
    "\" Sales - Clubs\" Or (ItemMovement.Description)=\"Transfers\"))"
Begin InputTables
    Name ="ItemMovement"
End
Begin OutputColumns
    Expression ="ItemMovement.Code"
    Alias ="CurrMMQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[CurrMM],[CurrMM]))"
    Alias ="JanQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Jan],[Jan]))"
    Alias ="FebQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Feb],[Feb]))"
    Alias ="MarQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Mar],[Mar]))"
    Alias ="AprQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Apr],[Apr]))"
    Alias ="MayQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[May],[May]))"
    Alias ="JunQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Jun],[Jun]))"
    Alias ="JulQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Jul],[Jul]))"
    Alias ="AugQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Aug],[Aug]))"
    Alias ="SepQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Sep],[Sep]))"
    Alias ="OctQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Oct],[Oct]))"
    Alias ="NovQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Nov],[Nov]))"
    Alias ="DecQty"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[Dec],[Dec]))"
    Alias ="YtdQ"
    Expression ="Sum(IIf([Description]=\"Transfers\",0-[YTDQty],[YTDQty]))"
End
Begin OrderBy
    Expression ="ItemMovement.Code"
    Flag =0
End
Begin Groups
    Expression ="ItemMovement.Code"
    GroupLevel =0
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
Begin
    Begin
        dbText "Name" ="JanQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FebQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MarQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AprQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MayQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="JunQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="JulQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AugQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SepQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OctQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NovQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrMMQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YtdQ"
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
    Bottom =116
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =35
        Top =17
        Right =198
        Bottom =131
        Top =0
        Name ="ItemMovement"
        Name =""
    End
End
