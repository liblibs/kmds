Operation =1
Option =0
Having ="(((IIf([qryEOSalesByDepartmentNoOS].[Transtype]='C' Or [qryEOSalesByDepartmentNo"
    "OS].[Transtype]='D',[qryEOSalesByDepartmentNoOS].[GroupDesc] & \" Bars\",[qryEOS"
    "alesByDepartmentNoOS].[GroupDesc]))<>\"Drops and Loans\"))"
Begin InputTables
    Name ="qryEOSalesByDepartmentNoOS"
    Name ="Agents"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="Agents.eMail"
    Expression ="Sysfile.AgentNumber"
    Expression ="Sysfile.Name"
    Expression ="qryEOSalesByDepartmentNoOS.GroupLvl"
    Expression ="qryEOSalesByDepartmentNoOS.Group"
    Alias ="GroupDescription"
    Expression ="IIf([qryEOSalesByDepartmentNoOS].[Transtype]='C' Or [qryEOSalesByDepartmentNoOS]"
        ".[Transtype]='D',[qryEOSalesByDepartmentNoOS].[GroupDesc] & \" Bars\",[qryEOSale"
        "sByDepartmentNoOS].[GroupDesc])"
    Alias ="SumOfTotalQ"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.TotalQ)"
    Alias ="TotalP"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.TotalP)"
    Alias ="SumOfQ1"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.Q1)"
    Alias ="SumOfD1"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.D1)"
    Alias ="SumOfQ2"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.Q2)"
    Alias ="SumOfD2"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.D2)"
    Alias ="SumOfQ3"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.Q3)"
    Alias ="SumOfD3"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.D3)"
    Alias ="SumOfQ4"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.Q4)"
    Alias ="SumOfD4"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.D4)"
    Alias ="SumOfTotalC"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.TotalC)"
    Alias ="SumOfTotalGP"
    Expression ="Sum(qryEOSalesByDepartmentNoOS.TotalGP)"
End
Begin Joins
    LeftTable ="Agents"
    RightTable ="Sysfile"
    Expression ="Agents.AgentNumber = Sysfile.AgentNumber"
    Flag =1
End
Begin Groups
    Expression ="Agents.eMail"
    GroupLevel =0
    Expression ="Sysfile.AgentNumber"
    GroupLevel =0
    Expression ="Sysfile.Name"
    GroupLevel =0
    Expression ="qryEOSalesByDepartmentNoOS.GroupLvl"
    GroupLevel =0
    Expression ="qryEOSalesByDepartmentNoOS.Group"
    GroupLevel =0
    Expression ="IIf([qryEOSalesByDepartmentNoOS].[Transtype]='C' Or [qryEOSalesByDepartmentNoOS]"
        ".[Transtype]='D',[qryEOSalesByDepartmentNoOS].[GroupDesc] & \" Bars\",[qryEOSale"
        "sByDepartmentNoOS].[GroupDesc])"
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
        dbText "Name" ="Sysfile.Name"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryEOSalesByDepartmentNoOS.GroupLvl"
        dbInteger "ColumnWidth" ="285"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfTotalQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQ1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfD1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQ2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfD2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQ3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfD3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQ4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfD4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfTotalC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfTotalGP"
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
    Right =1193
    Bottom =415
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =521
        Top =6
        Right =783
        Bottom =120
        Top =0
        Name ="qryEOSalesByDepartmentNoOS"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =174
        Bottom =255
        Top =0
        Name ="Agents"
        Name =""
    End
    Begin
        Left =224
        Top =35
        Right =367
        Bottom =284
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
