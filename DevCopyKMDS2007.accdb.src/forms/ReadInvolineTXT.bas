Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =1710
    Top =2670
    Right =10695
    Bottom =7605
    RecSrcDt = Begin
        0x4103b05415aee240
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =660
                    Width =3360
                    Height =480
                    Name ="Text0"
                    DefaultValue ="\"C:\\KMDS\\Involine.txt\""

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =660
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="Text0:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =1620
                    Width =2580
                    Height =300
                    TabIndex =1
                    Name ="Command2"
                    Caption ="Command2"
                    OnClick ="[Event Procedure]"

                End
            End
        End
    End
End
CodeBehindForm
' See "ReadInvolineTXT.cls"
