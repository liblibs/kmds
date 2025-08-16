Operation =1
Option =0
Where ="(((Cashier.ID)=[Forms]![CashiersAndProgramsSwitchBoard]![cmbCashier]))"
Begin InputTables
    Name ="Switchboard Items"
    Name ="CashiersAndPrograms"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
    Expression ="CashiersAndPrograms.Type"
    Expression ="CashiersAndPrograms.Allowed"
    Expression ="[Switchboard Items].Argument"
    Expression ="[Switchboard Items].SwitchboardID"
    Expression ="[Switchboard Items].ItemNumber"
    Expression ="[Switchboard Items].ItemText"
End
Begin Joins
    LeftTable ="CashiersAndPrograms"
    RightTable ="Cashier"
    Expression ="CashiersAndPrograms.CashierID = Cashier.ID"
    Flag =1
    LeftTable ="Switchboard Items"
    RightTable ="CashiersAndPrograms"
    Expression ="[Switchboard Items].Argument = CashiersAndPrograms.ProgramName"
    Flag =1
End
Begin OrderBy
    Expression ="[Switchboard Items].SwitchboardID"
    Flag =0
    Expression ="[Switchboard Items].ItemNumber"
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
        dbText "Name" ="CashiersAndPrograms.Type"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashiersAndPrograms.Allowed"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Switchboard Items].Argument"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Switchboard Items].SwitchboardID"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =189
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =446
        Top =39
        Right =580
        Bottom =183
        Top =0
        Name ="Switchboard Items"
        Name =""
    End
    Begin
        Left =167
        Top =25
        Right =376
        Bottom =184
        Top =0
        Name ="CashiersAndPrograms"
        Name =""
    End
    Begin
        Left =15
        Top =23
        Right =111
        Bottom =137
        Top =0
        Name ="Cashier"
        Name =""
    End
End
