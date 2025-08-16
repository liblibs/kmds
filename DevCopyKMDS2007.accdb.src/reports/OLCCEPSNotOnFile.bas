Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =16
    Left =1845
    Top =3075
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x0f9f88853384e440
    End
    RecordSource ="qryOLCCEPSForOrdersNotOnFile"
    Caption ="OLCC EPS Not On File"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="OLCCItem"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =960
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =3540
                    Height =540
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label12"
                    Caption ="OLCC EPS Not On File"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6060
                    Top =60
                    Width =3900
                    Height =540
                    ColumnOrder =0
                    FontSize =20
                    ForeColor =8355711
                    Name ="OrderNumber"
                    ControlSource ="OrderNumber"
                    GridlineColor =10921638

                    LayoutCachedLeft =6060
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =600
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                    Begin
                        Begin Label
                            Left =3660
                            Top =60
                            Width =2370
                            Height =540
                            FontSize =20
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label15"
                            Caption ="OrderNumber"
                            GridlineColor =10921638
                            LayoutCachedLeft =3660
                            LayoutCachedTop =60
                            LayoutCachedWidth =6030
                            LayoutCachedHeight =600
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =3
                    Left =360
                    Top =60
                    Width =1320
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="OLCCItem_Label"
                    Caption ="OLCCItem"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =1740
                    Top =60
                    Width =2700
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1740
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =1320
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Shipped_Label"
                    Caption ="Shipped"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =7260
                    Top =60
                    Width =1920
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Price_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7260
                    LayoutCachedTop =60
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =10080
                    Top =60
                    Width =1380
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="UnitsPerCase_Label"
                    Caption ="UnitsPerCase"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10080
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =1320
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="OLCCItem"
                    ControlSource ="OLCCItem"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =1740
                    Top =60
                    Width =2700
                    Height =330
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GridlineColor =10921638

                    LayoutCachedLeft =1740
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =60
                    Width =1320
                    Height =330
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Shipped"
                    ControlSource ="Shipped"
                    Format ="Standard"
                    GridlineColor =10921638

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7260
                    Top =60
                    Width =1920
                    Height =330
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"
                    GridlineColor =10921638

                    LayoutCachedLeft =7260
                    LayoutCachedTop =60
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =10440
                    Top =60
                    Width =1020
                    Height =330
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"
                    GridlineColor =10921638

                    LayoutCachedLeft =10440
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =390
                End
            End
        End
        Begin PageFooter
            Height =570
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text13"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Top =240
                    Width =5040
                    Height =330
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text14"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6420
                    LayoutCachedTop =240
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =570
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "OLCCEPSNotOnFile.cls"
