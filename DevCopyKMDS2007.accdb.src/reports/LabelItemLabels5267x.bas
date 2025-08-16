Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =24
    Width =2520
    DatasheetFontHeight =10
    ItemSuffix =5
    RecSrcDt = Begin
        0x2d44d62c89abe240
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
            Height =720
            OnPrint ="=LabelLayout(Reports!LabelItemLabels5267)"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =75
                    Top =540
                    Width =2445
                    Height =180
                    ColumnWidth =2355
                    FontSize =7
                    TabIndex =1
                    Name ="DescrTag"
                    ControlSource ="=[Descr] & \" - \" & [ID]"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =510
                    Top =120
                    Width =1368
                    Height =480
                    FontSize =24
                    Name ="Text2"
                    ControlSource ="UPCG"
                    FontName ="UPC-A Half Height"

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
