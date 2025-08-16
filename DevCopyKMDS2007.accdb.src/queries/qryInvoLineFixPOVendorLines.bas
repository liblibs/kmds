Operation =4
Option =0
Where ="(((InvoLine.SalesGroup)<>Left([InputCode],InStr([InputCode],\"-\")-1)) AND ((Inv"
    "oLine.TransType)=\"P\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Name ="InvoLine.SalesGroup"
    Expression ="Left([InputCode],InStr([InputCode],\"-\")-1)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
