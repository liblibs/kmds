Operation =1
Option =0
Begin InputTables
    Name ="qryADS_ItemMovement"
    Name ="Item"
    Name ="Vendor"
    Name ="AverageDaySupply"
End
Begin OutputColumns
    Expression ="qryADS_ItemMovement.Code"
    Alias ="WADSCurrMM"
    Expression ="[CurrMM]*IIf(IsNull([ADSCurrMM]),0,[ADSCurrMM])"
    Alias ="WADSJan"
    Expression ="[Jan]*IIf(IsNull([ADSJan]),0,[ADSJan])"
    Alias ="WADSFeb"
    Expression ="[Feb]*IIf(IsNull([ADSFeb]),0,[ADSFeb])"
    Alias ="WADSMar"
    Expression ="[Mar]*IIf(IsNull([ADSMar]),0,[ADSMar])"
    Alias ="WADSApr"
    Expression ="[Apr]*IIf(IsNull([ADSApr]),0,[ADSApr])"
    Alias ="WADSMay"
    Expression ="[May]*IIf(IsNull([ADSMay]),0,[ADSMay])"
    Alias ="WADSJun"
    Expression ="[Jun]*IIf(IsNull([ADSJun]),0,[ADSJun])"
    Alias ="WADSJul"
    Expression ="[Jul]*IIf(IsNull([ADSJul]),0,[ADSJul])"
    Alias ="WADSAug"
    Expression ="[Aug]*IIf(IsNull([ADSAug]),0,[ADSAug])"
    Alias ="WADSSep"
    Expression ="[Sep]*IIf(IsNull([ADSSep]),0,[ADSSep])"
    Alias ="WADSOct"
    Expression ="[Oct]*IIf(IsNull([ADSOct]),0,[ADSOct])"
    Alias ="WADSNov"
    Expression ="[Nov]*IIf(IsNull([ADSNov]),0,[ADSNov])"
    Alias ="WADSDec"
    Expression ="[Dec]*IIf(IsNull([ADSDec]),0,[ADSDec])"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Vendor"
    Expression ="Item.VendorID = Vendor.ID"
    Flag =1
    LeftTable ="Vendor"
    RightTable ="AverageDaySupply"
    Expression ="Vendor.AverageDaySupplyID = AverageDaySupply.ID"
    Flag =1
    LeftTable ="qryADS_ItemMovement"
    RightTable ="Item"
    Expression ="qryADS_ItemMovement.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "OrderBy" ="qryADS_ItemMovementWeightedByVendorADS.Code"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="WADSJan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSFeb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSMar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSApr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSMay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSJun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSJul"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSAug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSSep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSOct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSNov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSDec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WADSCurrMM"
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
    Bottom =110
    Left =0
    Top =8
    ColumnsShown =539
    Begin
        Left =49
        Top =6
        Right =304
        Bottom =120
        Top =0
        Name ="qryADS_ItemMovement"
        Name =""
    End
    Begin
        Left =365
        Top =9
        Right =461
        Bottom =123
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =514
        Top =4
        Right =610
        Bottom =118
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =682
        Top =1
        Right =844
        Bottom =115
        Top =0
        Name ="AverageDaySupply"
        Name =""
    End
End
