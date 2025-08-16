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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =136
    Left =585
    Top =1320
    OrderBy ="ID"
    RecSrcDt = Begin
        0x51124d4a9242e340
    End
    RecordSource ="qryMonthlySalesByDepartment"
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
            ControlSource ="Group"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Department"
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
                    Width =5205
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Monthly Sales By Department"
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
            Height =555
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =240
                    Top =180
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Items"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2340
                    Top =180
                    Width =600
                    Height =300
                    Name ="CurrMM_Label"
                    Caption ="Cur"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3060
                    Top =180
                    Width =540
                    Height =300
                    Name ="Jan_Label"
                    Caption ="Jan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3660
                    Top =180
                    Width =540
                    Height =300
                    Name ="Feb_Label"
                    Caption ="Feb"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4260
                    Top =180
                    Width =540
                    Height =300
                    Name ="Mar_Label"
                    Caption ="Mar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4860
                    Top =180
                    Width =540
                    Height =300
                    Name ="Apr_Label"
                    Caption ="Apr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5460
                    Top =180
                    Width =540
                    Height =300
                    Name ="May_Label"
                    Caption ="May"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6060
                    Top =180
                    Width =540
                    Height =300
                    Name ="Jun_Label"
                    Caption ="Jun"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6660
                    Top =180
                    Width =540
                    Height =300
                    Name ="Jul_Label"
                    Caption ="Jul"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7260
                    Top =180
                    Width =540
                    Height =300
                    Name ="Aug_Label"
                    Caption ="Aug"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7860
                    Top =180
                    Width =540
                    Height =300
                    Name ="Sep_Label"
                    Caption ="Sep"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8460
                    Top =180
                    Width =540
                    Height =300
                    Name ="Oct_Label"
                    Caption ="Oct"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9060
                    Top =180
                    Width =540
                    Height =300
                    Name ="Nov_Label"
                    Caption ="Nov"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9660
                    Top =180
                    Width =540
                    Height =300
                    Name ="Dec_Label"
                    Caption ="Dec"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =10200
                    Top =180
                    Width =540
                    Height =300
                    Name ="YTDQty_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =10200
                    LayoutCachedTop =180
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =480
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =540
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =420
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    BackStyle =1
                    Width =840
                    Height =420
                    FontSize =18
                    ForeColor =8388608
                    Name ="Group"
                    ControlSource ="Group"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    BackStyle =1
                    Left =720
                    Width =7119
                    Height =420
                    ColumnWidth =1185
                    FontSize =18
                    TabIndex =1
                    ForeColor =8388608
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =375
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =180
                    Width =672
                    Height =375
                    ColumnWidth =1260
                    FontSize =14
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =900
                    Width =2304
                    Height =375
                    ColumnWidth =1710
                    FontSize =14
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Times New Roman"

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
                    Left =60
                    Width =840
                    Height =225
                    ColumnWidth =570
                    FontSize =6
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10800
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3300
                    Width =480
                    Height =225
                    TabIndex =2
                    Name ="Jan"
                    ControlSource ="Jan"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3300
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3840
                    Width =480
                    Height =225
                    TabIndex =3
                    Name ="Feb"
                    ControlSource ="Feb"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3840
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4440
                    Width =480
                    Height =225
                    TabIndex =4
                    Name ="Mar"
                    ControlSource ="Mar"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4440
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5040
                    Width =480
                    Height =225
                    TabIndex =5
                    Name ="Apr"
                    ControlSource ="Apr"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5040
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5640
                    Width =480
                    Height =225
                    TabIndex =6
                    Name ="May"
                    ControlSource ="May"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6180
                    Width =540
                    Height =225
                    TabIndex =7
                    Name ="Jun"
                    ControlSource ="Jun"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6180
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Width =480
                    Height =225
                    TabIndex =8
                    Name ="Jul"
                    ControlSource ="Jul"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6840
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7440
                    Width =480
                    Height =225
                    TabIndex =9
                    Name ="Aug"
                    ControlSource ="Aug"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7440
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8040
                    Width =480
                    Height =225
                    TabIndex =10
                    Name ="Sep"
                    ControlSource ="Sep"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Width =480
                    Height =225
                    TabIndex =11
                    Name ="Oct"
                    ControlSource ="Oct"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8640
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Width =480
                    Height =225
                    TabIndex =12
                    Name ="Nov"
                    ControlSource ="Nov"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9240
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9780
                    Width =420
                    Height =225
                    TabIndex =13
                    Name ="Dec"
                    ControlSource ="Dec"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9780
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10260
                    Width =480
                    Height =225
                    TabIndex =14
                    Name ="YTDQty"
                    ControlSource ="Total"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =10260
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2700
                    Width =540
                    Height =225
                    TabIndex =15
                    Name ="Text80"
                    ControlSource ="CurrMM"
                    Format ="#,##0;(#,##0)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =900
                    Width =2100
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =60
                    Top =120
                    Width =1920
                    Height =225
                    Name ="Text35"
                    ControlSource ="DepartmentDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    Left =1980
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text73"
                    ControlSource ="=Count([Onhand])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3180
                    Top =120
                    Width =495
                    Height =225
                    TabIndex =2
                    Name ="Text82"
                    ControlSource ="=Sum([Jan])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =120
                    LayoutCachedWidth =3675
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3720
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =3
                    Name ="Text83"
                    ControlSource ="=Sum([Feb])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =120
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4320
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =4
                    Name ="Text85"
                    ControlSource ="=Sum([Mar])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =120
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =5
                    Name ="Text86"
                    ControlSource ="=Sum([Apr])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =120
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5520
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =6
                    Name ="Text87"
                    ControlSource ="=Sum([May])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =7
                    Name ="Text88"
                    ControlSource ="=Sum([Jun])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6720
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =8
                    Name ="Text89"
                    ControlSource ="=Sum([Jul])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =120
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7320
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =9
                    Name ="Text90"
                    ControlSource ="=Sum([Aug])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =10
                    Name ="Text91"
                    ControlSource ="=Sum([Sep])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =120
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8520
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =11
                    Name ="Text92"
                    ControlSource ="=Sum([Oct])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =120
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =12
                    Name ="Text93"
                    ControlSource ="=Sum([Nov])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =120
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9660
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =13
                    Name ="Text94"
                    ControlSource ="=Sum([Dec])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =120
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10140
                    Top =120
                    Width =615
                    Height =270
                    TabIndex =14
                    Name ="Text95"
                    ControlSource ="=Sum([Total])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =120
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2580
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =15
                    Name ="Text96"
                    ControlSource ="=Sum([CurrMM])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =120
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =390
                End
                Begin Line
                    Visible = NotDefault
                    Top =60
                    Width =10800
                    Name ="Line129"
                End
                Begin Line
                    Visible = NotDefault
                    Top =420
                    Width =10800
                    Name ="Line130"
                End
                Begin Line
                    Visible = NotDefault
                    Top =480
                    Width =10800
                    Name ="Line131"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =60
                    Top =180
                    Width =1200
                    Height =225
                    Name ="Text47"
                    ControlSource ="GroupDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    Left =1440
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text97"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =180
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3120
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =2
                    Name ="Text98"
                    ControlSource ="=Sum([Jan])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =120
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3720
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =3
                    Name ="Text99"
                    ControlSource ="=Sum([Feb])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =120
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4320
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =4
                    Name ="Text100"
                    ControlSource ="=Sum([Mar])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =120
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =5
                    Name ="Text101"
                    ControlSource ="=Sum([Apr])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =120
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5520
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =6
                    Name ="Text102"
                    ControlSource ="=Sum([May])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =7
                    Name ="Text103"
                    ControlSource ="=Sum([Jun])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6720
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =8
                    Name ="Text104"
                    ControlSource ="=Sum([Jul])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =120
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7320
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =9
                    Name ="Text105"
                    ControlSource ="=Sum([Aug])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =10
                    Name ="Text106"
                    ControlSource ="=Sum([Sep])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =120
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8520
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =11
                    Name ="Text107"
                    ControlSource ="=Sum([Oct])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =120
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =12
                    Name ="Text108"
                    ControlSource ="=Sum([Nov])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =120
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =13
                    Name ="Text109"
                    ControlSource ="=Sum([Dec])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =120
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10260
                    Top =120
                    Width =480
                    Height =270
                    TabIndex =14
                    Name ="Text110"
                    ControlSource ="=Sum([Total])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =120
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2520
                    Top =120
                    Width =540
                    Height =270
                    TabIndex =15
                    Name ="Text111"
                    ControlSource ="=Sum([CurrMM])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2520
                    LayoutCachedTop =120
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =390
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
            Height =495
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Label
                    Left =600
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                    LayoutCachedLeft =600
                    LayoutCachedWidth =1650
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    Left =1980
                    Width =720
                    Height =225
                    Name ="Text112"
                    ControlSource ="=Count([Onhand])"

                    LayoutCachedLeft =1980
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3120
                    Width =540
                    Height =270
                    TabIndex =1
                    Name ="Text113"
                    ControlSource ="=Sum([Jan])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3120
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3720
                    Width =540
                    Height =270
                    TabIndex =2
                    Name ="Text114"
                    ControlSource ="=Sum([Feb])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4320
                    Width =540
                    Height =270
                    TabIndex =3
                    Name ="Text115"
                    ControlSource ="=Sum([Mar])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Width =540
                    Height =270
                    TabIndex =4
                    Name ="Text116"
                    ControlSource ="=Sum([Apr])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4920
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5520
                    Width =540
                    Height =270
                    TabIndex =5
                    Name ="Text117"
                    ControlSource ="=Sum([May])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5520
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Width =540
                    Height =270
                    TabIndex =6
                    Name ="Text118"
                    ControlSource ="=Sum([Jun])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6120
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6720
                    Width =540
                    Height =270
                    TabIndex =7
                    Name ="Text119"
                    ControlSource ="=Sum([Jul])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6720
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7320
                    Width =540
                    Height =270
                    TabIndex =8
                    Name ="Text120"
                    ControlSource ="=Sum([Aug])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7320
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Width =540
                    Height =270
                    TabIndex =9
                    Name ="Text121"
                    ControlSource ="=Sum([Sep])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8520
                    Width =540
                    Height =270
                    TabIndex =10
                    Name ="Text122"
                    ControlSource ="=Sum([Oct])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8520
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Width =540
                    Height =270
                    TabIndex =11
                    Name ="Text123"
                    ControlSource ="=Sum([Nov])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9120
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9660
                    Width =540
                    Height =270
                    TabIndex =12
                    Name ="Text124"
                    ControlSource ="=Sum([Dec])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10140
                    Width =615
                    Height =270
                    TabIndex =13
                    Name ="Text125"
                    ControlSource ="=Sum([Total])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =10140
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2520
                    Width =540
                    Height =270
                    TabIndex =14
                    Name ="Text126"
                    ControlSource ="=Sum([CurrMM])"
                    Format ="#,##0;(#,##0)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =270
                End
                Begin Line
                    Visible = NotDefault
                    Width =10800
                    Name ="Line132"
                End
                Begin Line
                    Visible = NotDefault
                    Top =300
                    Width =10800
                    Name ="Line133"
                    LayoutCachedTop =300
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =300
                End
                Begin Line
                    Visible = NotDefault
                    Top =360
                    Width =10800
                    Name ="Line134"
                    LayoutCachedTop =360
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =360
                End
                Begin Line
                    Visible = NotDefault
                    Top =420
                    Width =10800
                    Name ="Line135"
                    LayoutCachedTop =420
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "50 Monthly Sales By Department.cls"
