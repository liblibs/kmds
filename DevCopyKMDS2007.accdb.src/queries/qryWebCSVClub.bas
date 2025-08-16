Operation =1
Option =0
Where ="(((Club.EMail)<>\"\"))"
Begin InputTables
    Name ="Club"
End
Begin OutputColumns
    Expression ="Club.ID"
    Expression ="Club.Name"
    Expression ="Club.Contact"
    Expression ="Club.Address1"
    Expression ="Club.Address2"
    Expression ="Club.City"
    Expression ="Club.St"
    Expression ="Club.Zip"
    Expression ="Club.Phone1"
    Expression ="Club.Phone2"
    Expression ="Club.Phone3"
    Expression ="Club.MaintainParLevels"
    Expression ="Club.OrderGuideByCategory"
    Expression ="Club.EMail"
    Expression ="Club.Licensee"
    Expression ="Club.SetupDate"
    Expression ="Club.ChangeDate"
    Expression ="Club.OuncesColumn"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Club.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.EMail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.OuncesColumn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Licensee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Phone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Phone2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Phone3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.MaintainParLevels"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.OrderGuideByCategory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Zip"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-567
    Top =131
    Right =943
    Bottom =924
    Left =-1
    Top =-1
    Right =1487
    Bottom =382
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =380
        Top =0
        Name ="Club"
        Name =""
    End
End
