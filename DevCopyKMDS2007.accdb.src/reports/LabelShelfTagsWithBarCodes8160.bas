Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =24
    Width =3780
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =1050
    Top =1860
    RecSrcDt = Begin
        0x440c130e54bae240
    End
    RecordSource ="qryItemLabelsWithUPCGlyph"
    OnOpen ="=LabelSetup()"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            AutoLabel = NotDefault
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            FontWeight =300
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin BreakLevel
            ControlSource ="Descr"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            OnFormat ="=LabelInitialize()"
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            OnPrint ="=LabelLayout(Reports!LabelShelfTagsWithBarCodes8160)"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =144
                    Top =60
                    Width =3462
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Text2"
                    ControlSource ="Descr"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =144
                    Top =480
                    Width =1587
                    Height =900
                    FontSize =26
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="UPCG"
                    FontName ="UPC-A"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1800
                    Top =480
                    Height =420
                    FontSize =16
                    FontWeight =700
                    TabIndex =2
                    Name ="Price"
                    ControlSource ="CurPrice"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3312
                    Top =480
                    Width =285
                    Height =360
                    FontSize =16
                    FontWeight =700
                    TabIndex =3
                    Name ="S"
                    ControlSource ="S"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2730
                    Top =900
                    Width =870
                    TabIndex =4
                    Name ="SalesDpt"
                    ControlSource ="SalesDpt"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1800
                    Top =900
                    Width =930
                    Height =480
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
                    Name ="ItemID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2730
                    Top =1140
                    Width =870
                    TabIndex =6
                    Name ="SizeD"
                    ControlSource ="SizeD"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
