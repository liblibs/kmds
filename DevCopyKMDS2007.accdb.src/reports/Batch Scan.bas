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
    Width =4617
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =1005
    Top =240
    RecSrcDt = Begin
        0x1f4b1f7a9cfce340
    End
    RecordSource ="qryScanBatchLabels"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            BorderLineStyle =0
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
            CanGrow = NotDefault
            Height =1080
            Name ="Detail"
            Begin
                Begin Rectangle
                    BorderWidth =2
                    Left =72
                    Width =3528
                    Height =900
                    Name ="Box4"
                    LayoutCachedLeft =72
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =150
                    Top =120
                    Width =1662
                    Height =660
                    FontSize =24
                    FontWeight =700
                    Name ="Text2"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =150
                    LayoutCachedTop =120
                    LayoutCachedWidth =1812
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2085
                    Top =240
                    Width =1452
                    Height =540
                    FontSize =26
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="UPCG"
                    FontName ="UPC-A Half Height"

                    LayoutCachedLeft =2085
                    LayoutCachedTop =240
                    LayoutCachedWidth =3537
                    LayoutCachedHeight =780
                End
            End
        End
        Begin PageFooter
            Height =420
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Width =4395
                    Height =255
                    FontSize =6
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                    LayoutCachedWidth =4395
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Top =240
                    Width =4395
                    Height =180
                    FontSize =6
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text6"
                    ControlSource ="=\"www.kmds.us\""
                    FontName ="Times New Roman"

                    LayoutCachedTop =240
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =420
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "Batch Scan.cls"
