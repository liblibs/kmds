Operation =1
Option =0
Begin InputTables
    Name ="CheckCashing"
End
Begin OutputColumns
    Expression ="CheckCashing.ID"
    Expression ="CheckCashing.FirstName"
    Expression ="CheckCashing.LastName"
    Expression ="CheckCashing.Address"
    Alias ="CityStateZ"
    Expression ="[City] & \", \" & [State] & \"  \" & [zip]"
    Expression ="CheckCashing.BankAccount"
    Expression ="CheckCashing.Zip"
End
Begin OrderBy
    Expression ="CheckCashing.FirstName"
    Flag =0
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
        dbText "Name" ="CityStateZ"
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
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =191
        Bottom =120
        Top =0
        Name ="CheckCashing"
        Name =""
    End
End
