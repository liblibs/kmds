Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Sysfile"
    Name ="InvoLine5"
    Name ="Item"
    Name ="Size"
End
Begin OutputColumns
    Expression ="Sysfile.AgentNumber"
    Expression ="Sysfile.Name"
    Expression ="Sysfile.Add1"
    Expression ="Sysfile.Add2"
    Expression ="Sysfile.Add3"
    Expression ="Sysfile.Phone"
    Alias ="InvoiceNumber"
    Expression ="[CashSysId] & \"-\" & [ShiftID] & \"-\" & [InvoHeadID]"
    Expression ="InvoLine5.SetupDate"
    Alias ="CustomerID"
    Expression ="0"
    Alias ="SKU"
    Expression ="InvoLine5.Code"
    Expression ="Item.Description"
    Alias ="SizeDesc"
    Expression ="Size.Description"
    Alias ="UPC"
    Expression ="IIf([transtype]=\"C\",[Code],[InputCode])"
    Alias ="UnitPrice"
    Expression ="InvoLine5.Price/[quantity]"
    Expression ="InvoLine5.Quantity"
    Alias ="ExtendPrice"
    Expression ="InvoLine5.Price"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Size"
    Expression ="InvoLine5.ID = Size.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine5.CashSysID"
    Flag =0
    Expression ="InvoLine5.InvoHeadID"
    Flag =0
    Expression ="InvoLine5.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Sysfile.Name"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Add1"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.SetupDate"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNumber"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPC"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeDesc"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitPrice"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustomerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SKU"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendPrice"
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
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =540
        Top =-2
        Right =636
        Bottom =112
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =417
        Top =0
        Right =513
        Bottom =114
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =668
        Top =0
        Right =764
        Bottom =114
        Top =0
        Name ="Size"
        Name =""
    End
End
