Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    RowHeight =525
    DatasheetFontHeight =11
    ItemSuffix =16
    Left =645
    Top =3540
    Right =15420
    Bottom =11025
    RecSrcDt = Begin
        0xd6bc6e428be7e440
    End
    RecordSource ="qryWebOrderTotalByInvoice"
    DatasheetFontName ="Calibri"
    AllowFormView =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =7665
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =600
                    Height =315
                    ColumnWidth =2130
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="mm/dd hh:nn am/pm"
                    GridlineColor =10921638
                    ShowDatePicker =0

                    LayoutCachedLeft =2100
                    LayoutCachedTop =600
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =600
                            Width =1095
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="Setup Date"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =600
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1020
                    Width =3300
                    Height =315
                    ColumnWidth =3285
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="WebOrderNotes"
                    ControlSource ="OrderNotes"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1020
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1335
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1020
                            Width =1605
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label2"
                            Caption ="Notes"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1440
                    Height =315
                    ColumnWidth =2595
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="LicenseeName"
                    ControlSource ="Customer"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1440
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label3"
                            Caption ="Customer"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =240
                    Top =3780
                    Width =7170
                    Height =3585
                    TabIndex =3
                    BorderColor =10921638
                    Name ="qryWebOrdersWithLines subform"
                    SourceObject ="Form.qryWebOrdersWithLines subform"
                    LinkChildFields ="InvoiceNo"
                    LinkMasterFields ="InvoiceNo"
                    EventProcPrefix ="qryWebOrdersWithLines_subform"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =3780
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =7365
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =240
                            Top =3540
                            Width =3195
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="qryWebOrdersWithLines subform Label"
                            Caption ="qryWebOrdersWithLines subform"
                            EventProcPrefix ="qryWebOrdersWithLines_subform_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =3540
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =3855
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4740
                    Top =2100
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="InvoiceNo"
                    ControlSource ="InvoiceNo"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2415
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2940
                            Top =2100
                            Width =1020
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label9"
                            Caption ="Invoice #"
                            GridlineColor =10921638
                            LayoutCachedLeft =2940
                            LayoutCachedTop =2100
                            LayoutCachedWidth =3960
                            LayoutCachedHeight =2415
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4740
                    Top =2520
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="InvQty"
                    ControlSource ="InvQty"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2835
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2940
                            Top =2520
                            Width =885
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label10"
                            Caption ="Quantity"
                            GridlineColor =10921638
                            LayoutCachedLeft =2940
                            LayoutCachedTop =2520
                            LayoutCachedWidth =3825
                            LayoutCachedHeight =2835
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4740
                    Top =2940
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="InvDollars"
                    ControlSource ="InvDollars"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =2940
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =3255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2940
                            Top =2940
                            Width =1005
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label11"
                            Caption ="Dollars"
                            GridlineColor =10921638
                            LayoutCachedLeft =2940
                            LayoutCachedTop =2940
                            LayoutCachedWidth =3945
                            LayoutCachedHeight =3255
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2100
                    Top =1875
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text12"
                    ControlSource ="eMail"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1875
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =2190
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1875
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label13"
                            Caption ="eMail"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =1875
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =2190
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =420
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Contact"
                    ControlSource ="Contact"
                    GridlineColor =10921638

                    LayoutCachedLeft =5700
                    LayoutCachedTop =420
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3900
                            Top =420
                            Width =780
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label14"
                            Caption ="Contact"
                            GridlineColor =10921638
                            LayoutCachedLeft =3900
                            LayoutCachedTop =420
                            LayoutCachedWidth =4680
                            LayoutCachedHeight =735
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7440
                    Top =1920
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Phone1"
                    ControlSource ="Phone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =7440
                    LayoutCachedTop =1920
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =2235
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5640
                            Top =1920
                            Width =780
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label15"
                            Caption ="Phone1"
                            GridlineColor =10921638
                            LayoutCachedLeft =5640
                            LayoutCachedTop =1920
                            LayoutCachedWidth =6420
                            LayoutCachedHeight =2235
                        End
                    End
                End
            End
        End
    End
End
