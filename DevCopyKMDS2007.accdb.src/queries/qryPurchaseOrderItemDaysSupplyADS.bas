Operation =1
Option =0
Where ="(((InvoLine.SetupDate) Between [parBeginDate] And [parEndDate]) AND ((InvoLine.T"
    "ransType)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"X\") AND"
    " ((Item.SalesDepartmentID) Between [parBeginSalesDepartmentID] And [parEndSalesD"
    "epartmentID]))"
Having ="(((Item.VendorID)=[parVendorId]))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Alias ="DateStart"
    Expression ="[parBeginDate]"
    Alias ="DateEnd"
    Expression ="[parEndDate]"
    Expression ="Item.VendorID"
    Alias ="ItemId"
    Expression ="InvoLine.Code"
    Alias ="Sold"
    Expression ="Sum(IIf([TransType]=\"X\",0-[Quantity],[Quantity]))"
    Alias ="SoldDays"
    Expression ="(DateDiff(\"d\",[parBeginDate],[parEndDate])+1)"
End
Begin Parameters
    Name ="parBeginDate"
    Flag =8
    Name ="parEndDate"
    Flag =8
    Name ="parVendorID"
    Flag =4
    Name ="parBeginSalesDepartmentID"
    Flag =4
    Name ="parEndSalesDepartmentID"
    Flag =4
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.Code"
    Flag =0
End
Begin Groups
    Expression ="Item.VendorID"
    GroupLevel =0
    Expression ="InvoLine.Code"
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
        dbText "Name" ="Sold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateStart"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateEnd"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemId"
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
    Bottom =266
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =213
        Bottom =255
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =251
        Top =6
        Right =426
        Bottom =255
        Top =0
        Name ="Item"
        Name =""
    End
End
