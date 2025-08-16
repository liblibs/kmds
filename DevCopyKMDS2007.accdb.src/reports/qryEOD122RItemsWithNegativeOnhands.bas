Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
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
    ItemSuffix =68
    Left =2985
    Top =690
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x637fd93d89b6e540
    End
    RecordSource ="qryEOItemsWithNegativeOnhands"
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
        Begin BreakLevel
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10695
                    Height =510
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =3038
                    Top =60
                    Width =4755
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="ItemsWithNegativeOnhands"
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
                    Top =1140
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =600
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =240
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4350
                    Top =120
                    Width =1365
                    Height =300
                    Name ="D2_Label"
                    Caption ="UPCCode"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =480
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5940
                    Top =120
                    Width =810
                    Height =300
                    Name ="Label55"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6825
                    Top =120
                    Width =555
                    Height =300
                    Name ="Label57"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =510
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =1680
                    Top =120
                    Width =963
                    Height =330
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="Group"
                    ControlSource ="Group"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =120
                    LayoutCachedWidth =2643
                    LayoutCachedHeight =450
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =60
                    Top =120
                    Width =1503
                    Height =330
                    ColumnWidth =1185
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =1563
                    LayoutCachedHeight =450
                End
                Begin Label
                    Visible = NotDefault
                    Left =3900
                    Top =120
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Width =714
                    Height =300
                    ColumnWidth =1260
                    FontSize =11
                    ForeColor =8388608
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Times New Roman"

                    LayoutCachedWidth =714
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1020
                    Width =2763
                    Height =300
                    ColumnWidth =1710
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =3783
                    LayoutCachedHeight =300
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =240
                    Width =1080
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1320
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1320
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    Left =4140
                    Width =1710
                    Height =225
                    ColumnWidth =735
                    TabIndex =2
                    Name ="DollarsSoldToday"
                    ControlSource ="UPCCode"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =6060
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6060
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =6900
                    Width =660
                    Height =225
                    TabIndex =4
                    Name ="Text58"
                    ControlSource ="Price"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6900
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Width =960
                    Height =225
                    TabIndex =5
                    Name ="Ext"
                    ControlSource ="Ext"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =5880
                    Top =120
                    Width =900
                    Height =225
                    FontWeight =700
                    Name ="Text61"
                    ControlSource ="=Sum([Onhand])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =120
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7740
                    Top =120
                    Width =1140
                    Height =225
                    FontWeight =700
                    TabIndex =1
                    Name ="Text62"
                    ControlSource ="=Sum([Ext])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =120
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =345
                    CurrencySymbol ="$"
                End
                Begin Line
                    Left =6000
                    Top =60
                    Width =2880
                    Name ="Line63"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =60
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    OverlapFlags =4
                    Left =2880
                    Top =60
                    Width =714
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="Text66"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3594
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =1
                    Left =3780
                    Top =60
                    Width =1983
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="Text67"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =60
                    LayoutCachedWidth =5763
                    LayoutCachedHeight =360
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter4"
        End
        Begin PageFooter
            Height =600
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =5040
                    Height =300
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
                    Height =300
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
                    Width =5040
                    Height =300
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
                    Height =300
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
                    Left =60
                    Top =300
                    Width =10680
                    Height =300
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
                    Left =60
                    Width =10680
                    Height =255
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "qryEOD122RItemsWithNegativeOnhands.cls"
