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
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =109
    Left =1710
    Top =195
    OrderBy ="ID"
    RecSrcDt = Begin
        0x3633bee0d2fee340
    End
    RecordSource ="qryItemsWithMostRecentScanBatch"
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
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CategoryDescr"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1335
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =7380
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory By Category By Last Scan Batch"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1320
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Top =660
                    Width =10561
                    Height =300
                    ColumnOrder =0
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboSalesDepartmentIDBegin & \" \" & F"
                        "orms!PrintInventoryReport!cboSalesDepartmentIDBegin.Column(1) & \" and \" & Form"
                        "s!PrintInventoryReport!cboSalesDepartmentIDEnd & \" \" & Forms!PrintInventoryRep"
                        "ort!cboSalesDepartmentIDEnd.Column(1)"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Top =960
                    Width =10561
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboCategoryIDBegin & \" and \" & Form"
                        "s!PrintInventoryReport!cboCategoryIDEnd"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =495
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =240
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3990
                    Top =120
                    Width =825
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    Left =5100
                    Top =120
                    Width =540
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    Left =5790
                    Top =120
                    Width =855
                    Height =300
                    Name ="D2_Label"
                    Caption ="Ext Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7290
                    Top =120
                    Width =615
                    Height =300
                    Name ="D3_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8130
                    Top =120
                    Width =1050
                    Height =300
                    Name ="D4_Label"
                    Caption =" Extended"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    Left =9480
                    Top =120
                    Width =1230
                    Height =330
                    Name ="Label84"
                    Caption ="Most Recent"
                    LayoutCachedLeft =9480
                    LayoutCachedTop =120
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =450
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =285
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    Height =225
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =840
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =3000
                    Height =225
                    TabIndex =1
                    Name ="SalesDptDescr"
                    ControlSource ="SalesDptDescr"

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =480
            BreakLevel =1
            Name ="GroupHeader1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =960
                    Top =120
                    Height =225
                    Name ="CategoryDescr"
                    ControlSource ="CategoryDescr"

                    LayoutCachedLeft =960
                    LayoutCachedTop =120
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =345
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =285
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =12
                    TextAlign =1
                    Width =1080
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedWidth =1080
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Left =1140
                    Width =2880
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    Left =4020
                    Width =720
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4020
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5700
                    Width =900
                    ColumnWidth =735
                    TabIndex =3
                    Name ="ExtCost"
                    ControlSource ="=[Onhand]*[Cost]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    Left =4800
                    Width =840
                    ColumnWidth =735
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4800
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =240
                End
                Begin Line
                    Visible = NotDefault
                    Left =240
                    Top =240
                    Width =10380
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8880
                    Width =1815
                    ColumnWidth =3330
                    TabIndex =5
                    Name ="MostRecentCount"
                    ControlSource ="MostRecentCount"
                    Format ="Short Date"
                    FontName ="Arial Narrow"
                    ConditionalFormat = Begin
                        0x0100000076000000010000000000000005000000000000000a00000001000000 ,
                        0x00000000ccc8c200000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4e006f007700280029002d0031003800300000000000
                    End

                    LayoutCachedLeft =8880
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000000000000ccc8c200090000004e00 ,
                        0x6f007700280029002d0031003800300000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6915
                    Width =885
                    TabIndex =6
                    Name ="Price"
                    ControlSource ="Price"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6915
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7860
                    Width =900
                    TabIndex =7
                    Name ="Text86"
                    ControlSource ="=[Onhand]*[price]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7860
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter0"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =3960
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text89"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3960
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5640
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text90"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text92"
                    ControlSource ="=Sum([Onhand]*[Price])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =2580
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text93"
                    ControlSource ="=Count([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Height =225
                    TabIndex =4
                    Name ="Text94"
                    ControlSource ="CategoryDescr"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            Name ="GroupFooter1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =3960
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text101"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3960
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5640
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text102"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7800
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text103"
                    ControlSource ="=Sum([Onhand]*[Price])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =2580
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text104"
                    ControlSource ="=Count([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =540
                    Top =60
                    Width =1920
                    Height =225
                    TabIndex =4
                    Name ="Text106"
                    ControlSource ="SalesDptDescr"

                    LayoutCachedLeft =540
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =285
                End
            End
        End
        Begin PageFooter
            Height =555
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
                    Width =5040
                    Height =255
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
                    Width =4260
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
                    Width =5040
                    Height =255
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
                    Left =120
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
            Height =750
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =3780
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =60
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5460
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=Sum([Onhand]*[Cost])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5460
                    LayoutCachedTop =60
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =360
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7620
                    Top =120
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text87"
                    ControlSource ="=Sum([Onhand]*[Price])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =120
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =2400
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text88"
                    ControlSource ="=Count([Onhand])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =60
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =240
                    Top =480
                    Width =7185
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label108"
                    Caption ="Note:  Items with 'Most Recent' grayed out have not been scanned in the last 180"
                        " days."
                    LayoutCachedLeft =240
                    LayoutCachedTop =480
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =750
                End
            End
        End
    End
End
CodeBehindForm
' See "68 Items By Category By Last Inventory Date.cls"
