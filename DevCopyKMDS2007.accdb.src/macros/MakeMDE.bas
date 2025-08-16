Version =196611
ColumnsShown =0
Begin
    Action ="RunSQL"
    Argument ="UPDATE [Switchboard Items] SET [Switchboard Items].Command = [Command]+1\015\012"
        "WHERE ((([Switchboard Items].Argument)=\"default\"));\015\012"
    Argument ="-1"
End
Begin
    Action ="RunCommand"
    Argument ="7"
End
