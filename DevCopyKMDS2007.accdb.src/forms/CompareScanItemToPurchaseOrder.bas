Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14760
    DatasheetFontHeight =11
    ItemSuffix =24
    Left =1995
    Top =840
    Right =17010
    Bottom =9885
    RecSrcDt = Begin
        0x9a7cb345eb19e440
    End
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
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
            Height =9060
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =600
                    Top =1260
                    Width =14100
                    Height =7620
                    BorderColor =10921638
                    Name ="CompareScanItemToPurchaseOrderSubForm"
                    SourceObject ="Form.CompareScanItemToPurchaseOrderSub"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =1260
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =8880
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =600
                            Top =3060
                            Width =4815
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="qryCompareScanItemToPurchaseOrder_05 subform Label"
                            Caption ="qryCompareScanItemToPurchaseOrder_05 subform"
                            EventProcPrefix ="qryCompareScanItemToPurchaseOrder_05_subform_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =600
                            LayoutCachedTop =3060
                            LayoutCachedWidth =5415
                            LayoutCachedHeight =3375
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =1080
                    Top =180
                    Width =4686
                    Height =720
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Frame3"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =1080
                    LayoutCachedTop =180
                    LayoutCachedWidth =5766
                    LayoutCachedHeight =900
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =1200
                            Top =60
                            Width =2595
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label4"
                            Caption ="Show shipped items where"
                            GridlineColor =10921638
                            LayoutCachedLeft =1200
                            LayoutCachedTop =60
                            LayoutCachedWidth =3795
                            LayoutCachedHeight =375
                            BackThemeColorIndex =-1
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1260
                            Top =508
                            OptionValue =1
                            BorderColor =10921638
                            Name ="Option6"
                            GridlineColor =10921638

                            LayoutCachedLeft =1260
                            LayoutCachedTop =508
                            LayoutCachedWidth =1520
                            LayoutCachedHeight =748
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1490
                                    Top =480
                                    Width =900
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label7"
                                    Caption ="All Items"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1490
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =2390
                                    LayoutCachedHeight =795
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =2754
                            Top =538
                            OptionValue =2
                            BorderColor =10921638
                            Name ="Option8"
                            GridlineColor =10921638

                            LayoutCachedLeft =2754
                            LayoutCachedTop =538
                            LayoutCachedWidth =3014
                            LayoutCachedHeight =778
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2984
                                    Top =480
                                    Width =990
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label9"
                                    Caption ="Full Cases"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2984
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =3974
                                    LayoutCachedHeight =795
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =4314
                            Top =538
                            OptionValue =3
                            BorderColor =10921638
                            Name ="Option10"
                            GridlineColor =10921638

                            LayoutCachedLeft =4314
                            LayoutCachedTop =538
                            LayoutCachedWidth =4574
                            LayoutCachedHeight =778
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =4544
                                    Top =480
                                    Width =825
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label11"
                                    Caption ="Repacks"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =4544
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =5369
                                    LayoutCachedHeight =795
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7140
                    Top =180
                    Width =1335
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text15"
                    ControlSource ="=[CompareScanItemToPurchaseOrderSubForm].[Form]![SumQuantityShipped]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =180
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =495
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5940
                            Top =180
                            Width =990
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="txtTotals"
                            Caption ="All Items"
                            GridlineColor =10921638
                            LayoutCachedLeft =5940
                            LayoutCachedTop =180
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =495
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8700
                    Top =180
                    Width =1335
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text17"
                    ControlSource ="=[CompareScanItemToPurchaseOrderSubForm].[Form]![SumQuantityScanned]"
                    GridlineColor =10921638

                    LayoutCachedLeft =8700
                    LayoutCachedTop =180
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =180
                    Width =1275
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text19"
                    ControlSource ="=[CompareScanItemToPurchaseOrderSubForm].[Form]![SumDiff]"
                    GridlineColor =10921638

                    LayoutCachedLeft =10260
                    LayoutCachedTop =180
                    LayoutCachedWidth =11535
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7140
                    Top =615
                    Width =1335
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text20"
                    ControlSource ="=[CompareScanItemToPurchaseOrderSubForm].[Form]![SumExtShipped]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GridlineColor =10921638

                    LayoutCachedLeft =7140
                    LayoutCachedTop =615
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =930
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8700
                    Top =615
                    Width =1335
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text22"
                    ControlSource ="=[CompareScanItemToPurchaseOrderSubForm].[Form]![SumExtScanned]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GridlineColor =10921638

                    LayoutCachedLeft =8700
                    LayoutCachedTop =615
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =930
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =615
                    Width =1275
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text23"
                    ControlSource ="=[CompareScanItemToPurchaseOrderSubForm].[Form]![SumExtDiff]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10260
                    LayoutCachedTop =615
                    LayoutCachedWidth =11535
                    LayoutCachedHeight =930
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =11700
                    Top =240
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =8
                    Name ="cmdClose"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =11700
                    LayoutCachedTop =240
                    LayoutCachedWidth =12366
                    LayoutCachedHeight =846
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    UseTheme =0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =12420
                            Top =240
                            Width =585
                            Height =405
                            FontSize =8
                            Name ="Label65"
                            Caption ="&Close Form"
                            FontName ="Tahoma"
                            LayoutCachedLeft =12420
                            LayoutCachedTop =240
                            LayoutCachedWidth =13005
                            LayoutCachedHeight =645
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CompareScanItemToPurchaseOrder.cls"
