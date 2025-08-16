Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10695
    DatasheetFontHeight =10
    ItemSuffix =102
    Left =585
    Top =1320
    OrderBy ="SetupDate"
    RecSrcDt = Begin
        0x01b5a1114cb8e240
    End
    RecordSource ="qryItemMovementReport"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1740
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =420
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10620
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Item  Movement Report"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1140
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboItemIDBegin & \" and \" & Forms!Pr"
                        "intInventoryReport!cboItemIDEnd"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1440
                    Width =10621
                    Height =300
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text79"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =510
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =120
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5340
                    Top =120
                    Width =1140
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Purchases"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9120
                    Top =120
                    Width =1260
                    Height =300
                    Name ="D2_Label"
                    Caption ="Adjustments"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10635
                    Height =30
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10635
                    Height =30
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10635
                    Height =30
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10635
                    Height =30
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6540
                    Top =120
                    Width =780
                    Height =300
                    Name ="Label86"
                    Caption ="Counter"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7380
                    Top =120
                    Width =720
                    Height =300
                    Name ="Label87"
                    Caption ="Clubs"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8160
                    Top =120
                    Width =960
                    Height =300
                    Name ="Label88"
                    Caption ="Transfers"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =120
                    Top =60
                    Width =1140
                    Height =225
                    FontWeight =700
                    Name ="Group"
                    ControlSource ="ID"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =1
                    Left =1320
                    Top =60
                    Width =7119
                    Height =225
                    ColumnWidth =1185
                    FontWeight =700
                    TabIndex =1
                    Name ="GroupDesc"
                    ControlSource ="Description"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =8439
                    LayoutCachedHeight =285
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =300
                    Width =2100
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="SetupDate"
                    Format ="General Date"

                    LayoutCachedLeft =300
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =1
                    BackStyle =1
                    Left =2460
                    Width =1740
                    Height =225
                    ColumnWidth =735
                    TabIndex =1
                    Name ="ExtCost"
                    ControlSource ="InputCode"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000012010000030000000100000000000000000000001000000001000000 ,
                        0x00000000ccffcc00010000000000000011000000340000000100000000000000 ,
                        0xffff9900010000000000000035000000580000000100000000000000ff99cc00 ,
                        0x5b005400720061006e00730054007900700065005d003d002200430022000000 ,
                        0x00005b005400720061006e00730054007900700065005d003d00220050002200 ,
                        0x20004f00720020005b005400720061006e00730054007900700065005d003d00 ,
                        0x220058002200000000005b005400720061006e00730054007900700065005d00 ,
                        0x3d0022004100220020004f00720020005b005400720061006e00730054007900 ,
                        0x700065005d003d0022005200220000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ccffcc000f0000005b00 ,
                        0x5400720061006e00730054007900700065005d003d0022004300220000000000 ,
                        0x0000000000000000000000000000000000010000000000000001000000000000 ,
                        0x00ffff9900220000005b005400720061006e00730054007900700065005d003d ,
                        0x0022005000220020004f00720020005b005400720061006e0073005400790070 ,
                        0x0065005d003d0022005800220000000000000000000000000000000000000000 ,
                        0x000001000000000000000100000000000000ff99cc00220000005b0054007200 ,
                        0x61006e00730054007900700065005d003d0022004100220020004f0072002000 ,
                        0x5b005400720061006e00730054007900700065005d003d002200520022000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4260
                    Height =225
                    TabIndex =2
                    Name ="Emp"
                    ControlSource ="Emp"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Purchases"
                    ControlSource ="Purchases"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Counter"
                    ControlSource ="Counter"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Clubs"
                    ControlSource ="Clubs"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8400
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Transfers"
                    ControlSource ="Transfers"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9660
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="Adjustments"
                    ControlSource ="Adjustments"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin Line
                    Visible = NotDefault
                    Left =4320
                    Top =240
                    Width =6060
                    Name ="Line5"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10500
                    Width =180
                    TabIndex =8
                    Name ="TransType"
                    ControlSource ="TransType"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2700
                    Top =60
                    Width =2880
                    Height =225
                    Name ="Text47"
                    ControlSource ="Description"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =1
                    Name ="Text90"
                    ControlSource ="=Sum([Purchases])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =60
                    Width =720
                    TabIndex =2
                    Name ="Text91"
                    ControlSource ="=Sum([Counter])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =720
                    TabIndex =3
                    Name ="Text92"
                    ControlSource ="=Sum([Clubs])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8400
                    Top =60
                    Width =720
                    TabIndex =4
                    Name ="Text93"
                    ControlSource ="=Sum([Transfers])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9660
                    Top =60
                    Width =720
                    TabIndex =5
                    Name ="Text94"
                    ControlSource ="=Sum([Adjustments])"
                    Format ="#,###;-#;\"\";\"\""

                End
            End
        End
        Begin PageFooter
            Height =585
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =1920
                    Height =255
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =300
                    Width =4965
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =4380
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Width =4965
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =300
                    Width =10680
                    Height =255
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Width =10620
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =330
            Name ="ReportFooter"
            Begin
                Begin Label
                    Left =3960
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    Top =60
                    Width =720
                    Name ="Text95"
                    ControlSource ="=Sum([Purchases])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =60
                    Width =720
                    TabIndex =1
                    Name ="Text96"
                    ControlSource ="=Sum([Counter])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =720
                    TabIndex =2
                    Name ="Text97"
                    ControlSource ="=Sum([Clubs])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8400
                    Top =60
                    Width =720
                    TabIndex =3
                    Name ="Text98"
                    ControlSource ="=Sum([Transfers])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9660
                    Top =60
                    Width =720
                    TabIndex =4
                    Name ="Text99"
                    ControlSource ="=Sum([Adjustments])"
                    Format ="#,###;-#;\"\";\"\""

                End
            End
        End
    End
End
CodeBehindForm
' See "18 Item Movement.cls"
