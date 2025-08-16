Operation =1
Option =0
Begin InputTables
    Name ="Vendor"
    Name ="AverageDaySupply"
    Name ="AverageDaySupplyLevels"
End
Begin OutputColumns
    Alias ="VendorID"
    Expression ="Vendor.ID"
    Alias ="AverageDaySupplyID"
    Expression ="AverageDaySupply.ID"
    Alias ="AverageDaySupplyLevelsID"
    Expression ="AverageDaySupplyLevels.ID"
    Expression ="AverageDaySupplyLevels.DaysSupply"
    Expression ="AverageDaySupplyLevels.MinUnitSales"
    Expression ="AverageDaySupplyLevels.MaxUnitSales"
End
Begin Joins
    LeftTable ="Vendor"
    RightTable ="AverageDaySupply"
    Expression ="Vendor.AverageDaySupplyID=AverageDaySupply.ID"
    Flag =1
    LeftTable ="AverageDaySupply"
    RightTable ="AverageDaySupplyLevels"
    Expression ="AverageDaySupply.ID=AverageDaySupplyLevels.AverageDaySupplyID"
    Flag =1
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
        dbText "Name" ="VendorID"
        dbInteger "ColumnWidth" ="975"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageDaySupplyID"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageDaySupplyLevelsID"
        dbInteger "ColumnWidth" ="2535"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =4
    Top =229
    Right =1229
    Bottom =906
    Left =-1
    Top =-1
    Right =1187
    Bottom =444
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =22
        Top =6
        Right =233
        Bottom =435
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =353
        Top =12
        Right =593
        Bottom =216
        Top =0
        Name ="AverageDaySupply"
        Name =""
    End
    Begin
        Left =713
        Top =10
        Right =875
        Bottom =124
        Top =0
        Name ="AverageDaySupplyLevels"
        Name =""
    End
End
