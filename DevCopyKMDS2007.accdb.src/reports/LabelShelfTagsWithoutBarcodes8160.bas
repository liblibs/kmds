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
    ItemSuffix =6
    Left =525
    Top =1230
    RecSrcDt = Begin
        0x3a8b400954bae240
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
            OnPrint ="=LabelLayout(Reports!LabelShelfTagsWithoutBarcodes8160)"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =75
                    Top =120
                    Width =3597
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="Text2"
                    ControlSource ="Descr"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =210
                    Top =600
                    Width =2955
                    Height =465
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    Name ="Price"
                    ControlSource ="CurPrice"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3240
                    Top =660
                    Width =285
                    Height =360
                    FontSize =16
                    FontWeight =700
                    TabIndex =2
                    Name ="S"
                    ControlSource ="S"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =72
                    Top =1140
                    Width =1215
                    TabIndex =3
                    Name ="SalesDpt"
                    ControlSource ="SalesDpt"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =72
                    LayoutCachedTop =1140
                    LayoutCachedWidth =1287
                    LayoutCachedHeight =1380
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1368
                    Top =1140
                    Width =1185
                    FontWeight =400
                    TabIndex =4
                    Name ="ItemID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1368
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2553
                    LayoutCachedHeight =1380
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2310
                    Top =1140
                    Width =720
                    TabIndex =5
                    Name ="SizeD"
                    ControlSource ="SizeD"

                    LayoutCachedLeft =2310
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =1380
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3024
                    Top =1140
                    Width =645
                    TabIndex =6
                    Name ="Text5"
                    ControlSource ="POZ"

                    LayoutCachedLeft =3024
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3669
                    LayoutCachedHeight =1380
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
CodeBehindForm
' See "LabelShelfTagsWithoutBarcodes8160.cls"
