Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =70
    Top =600
    OrderBy ="ItemId"
    RecSrcDt = Begin
        0x0039124e0ab1e240
    End
    RecordSource ="PurchaseOrderItemDaysSupply"
    Caption ="qryPurchaseOrderItemDaysSupply"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =255
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2340
            Name ="ReportHeader"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10740
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =1
                    Left =120
                    Top =720
                    Width =900
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text56"
                    ControlSource ="VendorID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =1
                    Left =1080
                    Top =720
                    Width =9594
                    Height =510
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text57"
                    ControlSource ="VendorName"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    Left =120
                    Top =120
                    Width =8700
                    Height =510
                    FontSize =20
                    Name ="Label51"
                    Caption ="Purchase Order Create by Days Supply Worksheet"
                End
                Begin Label
                    Visible = NotDefault
                    Left =6840
                    Top =1980
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =1320
                    Width =10560
                    Height =600
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    Name ="txtDateStart"
                    ControlSource ="=\"Average Daily Sales (ADS) based on Sales from \" & Format([DateStart],\"mm/dd"
                        "/yyyy\") & \" thru \" & Format([DateEnd],\"mm/dd/yyyy\") & \" (Number of days = "
                        "\" & [SoldDays] & \")\""
                    Format ="Short Date"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =1980
                    Width =6000
                    Height =315
                    FontSize =11
                    FontWeight =700
                    TabIndex =3
                    Name ="txtDaysSupplyRequested"
                    ControlSource ="=\"Days Supply \" & [SupplyDays]"

                End
            End
        End
        Begin PageHeader
            Height =720
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5565
                    Top =120
                    Width =780
                    Height =540
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text58"
                    ControlSource ="=[SupplyDays] & \" Day Supply\""
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =6480
                            Top =120
                            Width =855
                            Height =555
                            Name ="Label67"
                            Caption ="Current OrderTo"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    Left =300
                    Top =360
                    Width =540
                    Height =300
                    Name ="Code_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =960
                    Top =360
                    Width =1125
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3180
                    Top =360
                    Width =885
                    Height =300
                    Name ="SizeDesc_Label"
                    Caption ="SizeDesc"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4305
                    Top =360
                    Width =480
                    Height =300
                    Name ="Sold_Label"
                    Caption ="Sold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4965
                    Top =360
                    Width =495
                    Height =300
                    Name ="AverageDailySales_Label"
                    Caption ="ADS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7740
                    Top =360
                    Width =840
                    Height =300
                    Name ="Onhand_Label"
                    Caption ="Avail*"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8625
                    Top =120
                    Width =555
                    Height =555
                    Name ="SOQRaw_Label"
                    Caption ="SOQ Raw"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9240
                    Top =105
                    Width =840
                    Height =555
                    Name ="OrderMultiple_Label"
                    Caption ="Order Multiple"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =10170
                    Top =360
                    Width =510
                    Height =300
                    Name ="SOQ_Label"
                    Caption ="SOQ"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10740
                    Name ="Line54"
                End
                Begin Line
                    BorderWidth =2
                    Top =660
                    Width =10740
                    Name ="Line55"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =180
                    Width =660
                    Height =225
                    Name ="Code"
                    ControlSource ="ItemId"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =900
                    Width =2430
                    Height =225
                    ColumnWidth =2700
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3405
                    Width =660
                    Height =225
                    TabIndex =2
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4125
                    Width =660
                    Height =225
                    TabIndex =3
                    Name ="Sold"
                    ControlSource ="Sold"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4845
                    Width =660
                    Height =225
                    TabIndex =4
                    Name ="AverageDailySales"
                    ControlSource ="AverageDailySales"
                    Format ="Standard"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5685
                    Width =660
                    Height =225
                    ColumnWidth =1725
                    TabIndex =5
                    Name ="RequiredOnhands"
                    ControlSource ="SDOrderTo"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Width =660
                    Height =225
                    ColumnWidth =855
                    TabIndex =6
                    Name ="Onhand"
                    ControlSource ="Avail"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Width =540
                    Height =225
                    ColumnWidth =990
                    TabIndex =7
                    Name ="SOQRaw"
                    ControlSource ="SOQRaw"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9435
                    Width =660
                    Height =225
                    TabIndex =8
                    Name ="OrderMultiple"
                    ControlSource ="OrderMultiple"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10020
                    Width =660
                    Height =225
                    TabIndex =9
                    Name ="SOQ"
                    ControlSource ="SOQ"
                    Format ="#,##0;-#,##0;\"\";\"Null\015\012\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6480
                    Width =840
                    Height =225
                    TabIndex =10
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7440
                    Width =300
                    Height =225
                    TabIndex =11
                    Name ="Text68"
                    ControlSource ="=[SDOrderTo]-[OrderTo]"
                    Format ="General Number"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0xff000000ffffff00000000000600000003000000050000000100000000800000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x010002000000000000000500000001000000ff000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000006000000010000 ,
                        0x0000800000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin Line
                    Visible = NotDefault
                    Left =180
                    Top =240
                    Width =10380
                    Name ="Line5"
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
                    Width =4905
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
                    TextFontFamily =18
                    Left =60
                    Width =1920
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
                    Width =4905
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
                    Width =10620
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
                    Left =60
                    Width =10545
                    Height =285
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
            Height =840
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextAlign =3
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label39"
                    Caption ="Grand Total"
                End
                Begin Label
                    TextAlign =3
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label42"
                    Caption ="Grand Total"
                End
                Begin Label
                    TextAlign =3
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label45"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9780
                    Width =960
                    Height =300
                    FontSize =10
                    Name ="SOQ Grand Total Sum"
                    ControlSource ="=Sum([SOQ])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="SOQ_Grand_Total_Sum"

                End
                Begin Label
                    TextAlign =3
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label48"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =1560
                    Top =360
                    Width =5940
                    Height =480
                    Name ="Label63"
                    Caption ="*Avail = (Onhand + Ordered) - (Reserved - Sold today)"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7980
                    Width =960
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text66"
                    ControlSource ="=Sum([SOQRaw])"
                    FontName ="Times New Roman"

                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrderItemDaysSupply.cls"
