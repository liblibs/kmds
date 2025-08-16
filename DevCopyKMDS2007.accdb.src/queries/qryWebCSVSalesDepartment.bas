Operation =1
Option =0
Begin InputTables
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.SalesGroupID"
    Expression ="SalesDepartment.Description"
    Expression ="SalesDepartment.Discounts"
    Expression ="SalesDepartment.AllowClubDiscounts"
    Expression ="SalesDepartment.SetupDate"
    Expression ="SalesDepartment.ChangeDate"
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
        dbText "Name" ="SalesDepartment.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.SalesGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Discounts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.AllowClubDiscounts"
        dbInteger "ColumnWidth" ="3480"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-9
    Top =-12
    Right =1515
    Bottom =651
    Left =-1
    Top =-1
    Right =1501
    Bottom =171
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =449
        Top =16
        Right =754
        Bottom =248
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
