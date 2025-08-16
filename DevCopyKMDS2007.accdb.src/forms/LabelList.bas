Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10395
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =120
    Top =15
    Right =10770
    Bottom =9795
    RecSrcDt = Begin
        0xddf6b404333de440
    End
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin OptionButton
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =315
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6075
                    Width =4320
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    Name ="Auto_Date"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =6075
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =9480
            BackColor =-2147483607
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1500
                    Top =360
                    Width =2820
                    Height =315
                    BackColor =-2147483605
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="cboCategoryFrom"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [SubCategory].[ID], [SubCategory].[Description] FROM SubCategory ORDER BY"
                        " [Description]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =360
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =675
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =360
                            Width =975
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Category_Label"
                            Caption ="Category"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =360
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =680
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4500
                    Top =360
                    Width =2820
                    Height =315
                    TabIndex =1
                    BackColor =-2147483605
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="cboCategoryTo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [SubCategory].[ID], [SubCategory].[Description] FROM SubCategory ORDER BY"
                        " [Description]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4500
                    LayoutCachedTop =360
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =675
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1740
                            Top =60
                            Width =5760
                            Height =275
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label3"
                            Caption ="From                                                         To"
                            GridlineColor =10921638
                            LayoutCachedLeft =1740
                            LayoutCachedTop =60
                            LayoutCachedWidth =7500
                            LayoutCachedHeight =335
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =240
                    Top =840
                    Width =9720
                    Height =8640
                    FontSize =24
                    TabIndex =2
                    BoundColumn =3
                    BackColor =-2147483606
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstLabels"
                    RowSourceType ="Table/Query"
                    RowSource ="qryLabelList"
                    ColumnWidths ="0;7200;1440;0"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =840
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =9480
                    BackThemeColorIndex =-1
                End
                Begin OptionButton
                    OverlapFlags =85
                    Left =8280
                    Top =420
                    TabIndex =3
                    BorderColor =10921638
                    Name ="optShelf"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =8280
                    LayoutCachedTop =420
                    LayoutCachedWidth =8540
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8510
                            Top =390
                            Width =930
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label12"
                            Caption ="Shelf Tag"
                            GridlineColor =10921638
                            LayoutCachedLeft =8510
                            LayoutCachedTop =390
                            LayoutCachedWidth =9440
                            LayoutCachedHeight =705
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483607
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "LabelList.cls"
