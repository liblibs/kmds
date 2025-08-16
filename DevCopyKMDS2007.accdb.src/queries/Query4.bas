Operation =1
Option =0
Where ="(((SalesDepartment.Description) Like \"RJ*\") AND ((InvoLine.SetupDate) Between "
    "#1/1/2007# And #1/1/2008#) AND ((InvoLine.TransType)=\"S\" Or (InvoLine.TransTyp"
    "e)=\"C\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="SalesDepartment.Description"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.InputCode"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.Description"
    Flag =0
    Expression ="Item.Description"
    Flag =0
    Expression ="InvoLine.SetupDate"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="((Query4.InputCode=\"rx by sk\"))"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbInteger "ColumnWidth" ="2295"
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
    Bottom =262
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =153
        Bottom =239
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =221
        Top =16
        Right =317
        Bottom =130
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =411
        Top =8
        Right =507
        Bottom =122
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
