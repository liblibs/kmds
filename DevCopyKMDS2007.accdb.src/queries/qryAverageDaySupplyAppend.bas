Operation =3
Name ="AverageDaySupplyLevels"
Option =0
Begin InputTables
    Name ="AverageDaySupplyLevelsTemplate"
End
Begin OutputColumns
    Name ="GroupLevel"
    Expression ="AverageDaySupplyLevelsTemplate.GroupLevel"
    Name ="DaysSupply"
    Expression ="AverageDaySupplyLevelsTemplate.DaysSupply"
    Name ="MinUnitSales"
    Expression ="AverageDaySupplyLevelsTemplate.MinUnitSales"
    Name ="MaxUnitSales"
    Expression ="AverageDaySupplyLevelsTemplate.MaxUnitSales"
    Alias ="ADSID"
    Name ="AverageDaySupplyID"
    Expression ="Forms!AverageDaySupply!ID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ADSID"
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
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =236
        Bottom =120
        Top =0
        Name ="AverageDaySupplyLevelsTemplate"
        Name =""
    End
End
