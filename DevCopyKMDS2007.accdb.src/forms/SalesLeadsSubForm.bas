Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    RowHeight =735
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =1410
    Top =1320
    Right =13935
    Bottom =7800
    RecSrcDt = Begin
        0x66f8c5a11739e440
    End
    RecordSource ="SalesLeads"
    Caption ="SalesLeads subform"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =3210
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =360
                    Height =315
                    ColumnWidth =1440
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =360
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =675
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Label"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2670
                    Top =780
                    Width =6154
                    Height =600
                    ColumnWidth =7095
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text"
                    ControlSource ="Text"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =780
                    LayoutCachedWidth =8824
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Text_Label"
                            Caption ="Text"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1110
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =2340
                    Width =1620
                    Height =330
                    ColumnWidth =1620
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =2670
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2340
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2340
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2670
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2460
                    Top =1530
                    TabIndex =3
                    BorderColor =10921638
                    Name ="SpecificItemOnly"
                    ControlSource ="SpecificItemOnly"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =1530
                    LayoutCachedWidth =2720
                    LayoutCachedHeight =1770
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2690
                            Top =1500
                            Width =840
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label13"
                            Caption =" SpecificItemOnly"
                            GridlineColor =10921638
                            LayoutCachedLeft =2690
                            LayoutCachedTop =1500
                            LayoutCachedWidth =3530
                            LayoutCachedHeight =1815
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "SalesLeadsSubForm.cls"
