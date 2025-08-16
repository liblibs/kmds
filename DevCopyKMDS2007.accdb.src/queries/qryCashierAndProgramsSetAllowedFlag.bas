Operation =4
Option =0
Where ="(((CashiersAndPrograms.CashierID)=[parCashierID]))"
Begin InputTables
    Name ="CashiersAndPrograms"
End
Begin OutputColumns
    Name ="CashiersAndPrograms.Allowed"
    Expression ="[parAllowed]"
End
Begin Parameters
    Name ="parCashierID"
    Flag =4
    Name ="parAllowed"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =192
        Bottom =120
        Top =0
        Name ="CashiersAndPrograms"
        Name =""
    End
End
