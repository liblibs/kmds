Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8820
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =3645
    Top =660
    Right =12405
    Bottom =8325
    RecSrcDt = Begin
        0xc2d5edf40e9be240
    End
    Caption =" KMDS Software © "
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =7680
            BackColor =-2147483633
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =420
                    Width =8400
                    Height =7080
                    Name ="txtStatus"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =165
                            Top =120
                            Width =8415
                            Height =240
                            BackColor =65280
                            Name ="Label2"
                            Caption ="End of Day"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "EndOfDay.cls"
