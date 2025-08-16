Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12480
    DatasheetFontHeight =10
    ItemSuffix =14
    RecSrcDt = Begin
        0x453e73d750a1e240
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
            Height =4860
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =420
                    Width =3420
                    Height =300
                    Name ="txtDataBaseName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =420
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="Text0:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1200
                    Width =10500
                    Height =360
                    TabIndex =1
                    Name ="txtOldConnectString"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =1200
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="Text2:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1740
                    Width =10560
                    Height =420
                    TabIndex =2
                    Name ="txtNewConnectString"
                    DefaultValue ="\"DSN=KMDSSQL;UID=sa;PWD=;APP=Microsoft Office 2003;WSID=SK2;DATABASE=KMDSSQL\""

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =1740
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="Text4:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =4020
                    Width =1320
                    Height =420
                    TabIndex =3
                    Name ="Command7"
                    Caption ="Command7"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =2880
                    Width =3600
                    TabIndex =4
                    Name ="DataBaseName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =2880
                            Width =555
                            Height =240
                            Name ="Label9"
                            Caption ="Text8:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =2400
                    Width =3600
                    TabIndex =5
                    Name ="txtDir"
                    DefaultValue ="\"C:\\KMDS\\Backup\\2004\\09\\24\\\""

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =480
                            Top =2400
                            Width =735
                            Height =240
                            Name ="Label11"
                            Caption ="Directory"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =3240
                    Width =10500
                    Height =480
                    FontSize =16
                    TabIndex =6
                    Name ="txtCurFile"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =3240
                            Width =555
                            Height =240
                            Name ="Label13"
                            Caption ="Text8:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SetNewConnectString.cls"
