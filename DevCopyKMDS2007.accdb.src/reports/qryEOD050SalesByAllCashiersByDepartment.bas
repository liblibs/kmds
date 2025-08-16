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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =109
    Left =885
    Top =435
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x94413e2ff01ce440
    End
    RecordSource ="qryEOSalesbyAllCashiersByDepartment"
    Caption ="qryEOSalesbyAllCashiersByDepartment"
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
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="SalesOrTender"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="GroupLevel"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Group"
        End
        Begin BreakLevel
            ControlSource ="Department"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10680
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
                    Left =810
                    Top =60
                    Width =9165
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Sales Totals by All Cashiers By Register Report"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line50"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =10800
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10800
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =1095
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1680
                    Width =540
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =1
                            TextAlign =2
                            Left =60
                            Width =1537
                            Height =510
                            FontSize =20
                            Name ="Label67"
                            Caption ="Cashier:"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2280
                    Width =3300
                    Height =510
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =3
                    Left =1020
                    Top =720
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3540
                    Top =720
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4740
                    Top =720
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6180
                    Top =720
                    Width =645
                    Height =300
                    Name ="D1_Label"
                    Caption =" Reg 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6180
                    LayoutCachedTop =720
                    LayoutCachedWidth =6825
                    LayoutCachedHeight =1020
                End
                Begin Label
                    TextAlign =3
                    Left =7140
                    Top =720
                    Width =585
                    Height =300
                    Name ="D2_Label"
                    Caption ="Reg 2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7140
                    LayoutCachedTop =720
                    LayoutCachedWidth =7725
                    LayoutCachedHeight =1020
                End
                Begin Label
                    TextAlign =3
                    Left =8160
                    Top =720
                    Width =585
                    Height =300
                    Name ="D3_Label"
                    Caption ="Reg 3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8160
                    LayoutCachedTop =720
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =1020
                End
                Begin Label
                    TextAlign =3
                    Left =9000
                    Top =720
                    Width =645
                    Height =300
                    Name ="D4_Label"
                    Caption =" Reg 4"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =720
                    LayoutCachedWidth =9645
                    LayoutCachedHeight =1020
                End
                Begin Line
                    BorderWidth =1
                    Left =1020
                    Top =750
                    Width =9720
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1020
                    Top =720
                    Width =9720
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1020
                    Top =1050
                    Width =9720
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1020
                    Top =1080
                    Width =9720
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9915
                    Top =720
                    Width =690
                    Height =300
                    Name ="Label108"
                    Caption =" Reg 5"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9915
                    LayoutCachedTop =720
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =1020
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =360
            BreakLevel =2
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1080
                    Name ="Text94"
                    ControlSource ="GroupLevel"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =3
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =1140
                    Width =840
                    Height =225
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =2160
                    Width =1740
                    Height =225
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    Left =3180
                    Width =1185
                    Height =225
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="TotalQ"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Width =1185
                    FontWeight =700
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="TotalP"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5880
                    Width =915
                    Height =225
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="D1"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5880
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6840
                    Width =915
                    Height =225
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="D2"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6840
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7800
                    Width =915
                    Height =225
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="D3"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7800
                    LayoutCachedWidth =8715
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8760
                    Width =915
                    Height =225
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="D4"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8760
                    LayoutCachedWidth =9675
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9780
                    Width =915
                    Height =225
                    TabIndex =8
                    Name ="Text103"
                    ControlSource ="D5"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9780
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =540
                    Top =60
                    Width =180
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="*"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3180
                    Top =60
                    Width =1185
                    Height =225
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TotalQ"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =4545
                    Top =60
                    Width =1185
                    Height =225
                    TabIndex =1
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TotalP"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =2
                    Name ="Sum Of D1"
                    ControlSource ="=Sum([D1])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D1"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6840
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =3
                    Name ="Sum Of D2"
                    ControlSource ="=Sum([D2])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D2"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =60
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7800
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =4
                    Name ="Sum Of D3"
                    ControlSource ="=Sum([D3])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D3"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8715
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8760
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =5
                    Name ="Sum Of D4"
                    ControlSource ="=Sum([D4])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D4"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =60
                    LayoutCachedWidth =9675
                    LayoutCachedHeight =285
                End
                Begin Line
                    Left =60
                    Width =10680
                    BorderColor =12632256
                    Name ="Line58"
                End
                Begin TextBox
                    Left =810
                    Top =60
                    Width =2460
                    Height =225
                    TabIndex =6
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9780
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =7
                    Name ="Text104"
                    ControlSource ="=Sum([D5])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =360
            BreakLevel =2
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    Left =615
                    Top =60
                    Width =2100
                    Name ="txtGroupLevelDesc"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            TextFontFamily =34
                            Top =60
                            Width =285
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label68"
                            Caption ="**"
                            FontName ="Arial Narrow"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =375
                    Top =60
                    Width =180
                    TabIndex =1
                    Name ="txtGroupLevel"
                    ControlSource ="GroupLevel"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3120
                    Top =60
                    Width =1185
                    Height =225
                    TabIndex =2
                    Name ="Text88"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4485
                    Top =60
                    Width =1185
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text89"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =4
                    Name ="Text90"
                    ControlSource ="=Sum([D1])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6840
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =5
                    Name ="Text91"
                    ControlSource ="=Sum([D2])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =60
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7800
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =6
                    Name ="Text92"
                    ControlSource ="=Sum([D3])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =60
                    LayoutCachedWidth =8715
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8760
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =7
                    Name ="Text93"
                    ControlSource ="=Sum([D4])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =60
                    LayoutCachedWidth =9675
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9780
                    Top =60
                    Width =915
                    Height =225
                    TabIndex =8
                    Name ="Text105"
                    ControlSource ="=Sum([D5])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =1200
                    Top =240
                    Width =1800
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text78"
                    ControlSource ="SalesOrTender"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =660
                            Top =240
                            Width =420
                            Height =270
                            Name ="Label79"
                            Caption ="**"
                            FontName ="Arial Narrow"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3120
                    Top =240
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="Text80"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4320
                    Top =240
                    Width =1380
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text81"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5820
                    Top =240
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="Text82"
                    ControlSource ="=Sum([D1])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =240
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6780
                    Top =240
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="Text83"
                    ControlSource ="=Sum([D2])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =240
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7740
                    Top =240
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =5
                    Name ="Text84"
                    ControlSource ="=Sum([D3])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =240
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8700
                    Top =240
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =6
                    Name ="Text85"
                    ControlSource ="=Sum([D4])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =240
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =540
                End
                Begin Line
                    Top =60
                    Width =10680
                    BorderColor =12632256
                    Name ="Line86"
                End
                Begin Line
                    Top =120
                    Width =10680
                    BorderColor =12632256
                    Name ="Line87"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9720
                    Top =240
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =7
                    Name ="Text106"
                    ControlSource ="=Sum([D5])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =240
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =540
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =1515
                    Width =2460
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text95"
                    ControlSource ="=\"Over/Short\""
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =900
                            Width =435
                            Height =300
                            Name ="Label96"
                            Caption ="***"
                            FontName ="Arial Narrow"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4485
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="Text98"
                    ControlSource ="=0-Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"Balanced!\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5820
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="Text99"
                    ControlSource ="=0-Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5820
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6780
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="Text100"
                    ControlSource ="=0-Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6780
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7740
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="Text101"
                    ControlSource ="=0-Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8700
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =5
                    Name ="Text102"
                    ControlSource ="=0-Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8700
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9720
                    Width =915
                    Height =300
                    FontWeight =700
                    TabIndex =6
                    Name ="Text107"
                    ControlSource ="=0-Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9720
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =300
                End
            End
        End
        Begin PageFooter
            Height =660
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =360
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
                    Left =5760
                    Top =360
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
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line60"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =60
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
                    Left =5760
                    Top =60
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
                    Left =120
                    Top =360
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
                    Left =120
                    Top =60
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
            Height =390
            Name ="ReportFooter"
            Begin
                Begin Label
                    Visible = NotDefault
                    Left =60
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
    End
End
CodeBehindForm
' See "qryEOD050SalesByAllCashiersByDepartment.cls"
