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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =115
    Left =4725
    Top =930
    OrderBy ="ID"
    RecSrcDt = Begin
        0xaaba822f4b70e340
    End
    RecordSource ="qryItemsByDaysSupplyReport"
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
            ControlSource ="SaleDpt"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2040
            OnFormat ="[Event Procedure]"
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
                    Left =45
                    Top =60
                    Width =10635
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items By Days Supply of Onhands"
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
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1200
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=\"Average Daily Sales based on Sales Between \" & Format(Forms!PrintInventoryRe"
                        "port!txtDateRangeFrom,\"mm/dd/yyyy\") & \" And \" & Format(Forms!PrintInventoryR"
                        "eport!txtDateRangeTo,\"mm/dd/yyyy\") & \" ( \" & [Days] & \" days)\""
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1500
                    Width =10621
                    Height =300
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text112"
                    ControlSource ="=\"Days Supply range is between \" & CLng(Forms!PrintInventoryReport!txtNumericR"
                        "angeFrom) & \" And \" & CLng(Forms!PrintInventoryReport!txtNumericRangeTo)"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =735
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
                    Left =3360
                    Top =120
                    Width =1140
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10620
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5820
                    Top =120
                    Width =840
                    Height =300
                    Name ="Label86"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6660
                    Top =120
                    Width =1020
                    Height =540
                    Name ="Label87"
                    Caption ="Avg Dly Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7740
                    Top =120
                    Width =900
                    Height =600
                    Name ="Label88"
                    Caption ="Days Supply"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =180
                    Top =720
                    Width =10560
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8820
                    Top =120
                    Width =780
                    Height =585
                    Name ="Label110"
                    Caption ="Order To"
                End
                Begin Label
                    TextAlign =3
                    Left =9720
                    Top =120
                    Width =900
                    Height =600
                    Name ="Label111"
                    Caption ="30 Days Supply"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4620
                    Top =120
                    Width =1020
                    Height =540
                    Name ="Label113"
                    Caption ="Sold"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    ColumnWidth =615
                    Name ="SaleDpt"
                    ControlSource ="SaleDpt"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =3060
                    ColumnWidth =2730
                    TabIndex =1
                    Name ="SaleDptDesc"
                    ControlSource ="SaleDptDesc"

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
                    Left =1320
                    Width =2334
                    ColumnWidth =1185
                    FontWeight =700
                    Name ="GroupDesc"
                    ControlSource ="Description"

                    LayoutCachedLeft =1320
                    LayoutCachedWidth =3654
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    Width =1260
                    FontWeight =700
                    TabIndex =1
                    Name ="Group"
                    ControlSource ="ID"

                    LayoutCachedWidth =1260
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3600
                    Width =720
                    Height =225
                    ColumnWidth =885
                    TabIndex =2
                    Name ="Price"
                    ControlSource ="Price"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5700
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Onhand"
                    ControlSource ="Onhand"

                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6900
                    Width =795
                    Height =225
                    ColumnWidth =1770
                    TabIndex =4
                    Name ="AverageDailySales"
                    ControlSource ="AverageDailySales"
                    Format ="Standard"

                    LayoutCachedLeft =6900
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Width =720
                    Height =225
                    ColumnWidth =1635
                    TabIndex =5
                    Name ="DaysSupply"
                    ControlSource ="DaysSupply"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8760
                    Width =780
                    Height =225
                    TabIndex =6
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9840
                    Width =780
                    Height =225
                    TabIndex =7
                    Name ="Text109"
                    ControlSource ="=[AverageDailySales]*30"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Width =720
                    Height =225
                    TabIndex =8
                    Name ="Text114"
                    ControlSource ="Sold"

                End
                Begin Line
                    Visible = NotDefault
                    Left =3600
                    Top =240
                    Width =7080
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =180
            Name ="GroupFooter4"
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "10 Items By Days Supply of Onhand.cls"
