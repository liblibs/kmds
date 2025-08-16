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
    Width =10681
    DatasheetFontHeight =10
    ItemSuffix =86
    Left =315
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x848d6252ec1ce440
    End
    RecordSource ="qryEOSalesByDepartmentNoOS"
    Caption ="qryEOD040SalesByDepartment1"
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
            ControlSource ="GroupLvl"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Group"
        End
        Begin BreakLevel
            ControlSource ="Department"
        End
        Begin BreakLevel
            ControlSource ="DepartmentDesc"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
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
                    Width =10620
                    Height =540
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
                    Left =60
                    Top =60
                    Width =10620
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Sales Totals by Register Report"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10680
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
                Begin Subform
                    Left =8760
                    Top =240
                    Width =1816
                    Height =720
                    TabIndex =1
                    Name ="qryEOS040CustomerCount"
                    SourceObject ="Report.qryEOS040CustomerCount"

                End
            End
        End
        Begin PageHeader
            Height =900
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =720
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =720
                    LayoutCachedTop =120
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =3375
                    Top =120
                    Width =885
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3375
                    LayoutCachedTop =120
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =4650
                    Top =120
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4650
                    LayoutCachedTop =120
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =5700
                    Top =120
                    Width =645
                    Height =300
                    Name ="D1_Label"
                    Caption =" Reg 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5700
                    LayoutCachedTop =120
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =6945
                    Top =120
                    Width =585
                    Height =300
                    Name ="D2_Label"
                    Caption ="Reg 2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6945
                    LayoutCachedTop =120
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =8145
                    Top =120
                    Width =585
                    Height =300
                    Name ="D3_Label"
                    Caption ="Reg 3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8145
                    LayoutCachedTop =120
                    LayoutCachedWidth =8730
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =9255
                    Top =120
                    Width =525
                    Height =330
                    Name ="D4_Label"
                    Caption =" Reg 4"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9255
                    LayoutCachedTop =120
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =450
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1080
                    Top =120
                    Width =9600
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1080
                    Top =60
                    Width =9600
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =480
                    Width =9600
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =540
                    Width =9600
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9975
                    Top =120
                    Width =690
                    Height =330
                    Name ="Label82"
                    Caption =" Reg 5"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9975
                    LayoutCachedTop =120
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =450
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =540
                    Width =840
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =540
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Left =1560
                    Width =1740
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    Left =2835
                    Width =1185
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="TotalQ"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2835
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3945
                    Width =1185
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="TotalP"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3945
                    LayoutCachedWidth =5130
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5160
                    Width =1185
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="D1"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5160
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6345
                    Width =1185
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="D2"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6345
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7545
                    Width =1185
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="D3"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7545
                    LayoutCachedWidth =8730
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8715
                    Width =1065
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="D4"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8715
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Width =1065
                    TabIndex =8
                    Name ="D5"
                    ControlSource ="D5"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =240
                    Top =150
                    Width =450
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="*"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2820
                    Top =120
                    Width =1185
                    Height =270
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TotalQ"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =120
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3930
                    Top =120
                    Width =1185
                    Height =270
                    FontWeight =700
                    TabIndex =1
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TotalP"

                    LayoutCachedLeft =3930
                    LayoutCachedTop =120
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5160
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="Sum Of D1"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D1"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =120
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6345
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="Sum Of D2"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D2"

                    LayoutCachedLeft =6345
                    LayoutCachedTop =120
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7545
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="Sum Of D3"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D3"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =120
                    LayoutCachedWidth =8730
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8715
                    Top =120
                    Width =1065
                    Height =300
                    TabIndex =5
                    Name ="Sum Of D4"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D4"

                    LayoutCachedLeft =8715
                    LayoutCachedTop =120
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =540
                    Top =60
                    Width =10140
                    BorderColor =12632256
                    Name ="Line58"
                End
                Begin TextBox
                    Left =780
                    Top =150
                    Width =2460
                    TabIndex =6
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Top =120
                    Width =1065
                    Height =300
                    TabIndex =7
                    Name ="Sum Of D5"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D5"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =120
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =420
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2820
                    Top =60
                    Width =1185
                    Height =300
                    Name ="Text69"
                    ControlSource ="=Sum([TotalQ])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =60
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3930
                    Top =60
                    Width =1185
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3930
                    LayoutCachedTop =60
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5160
                    Top =60
                    Width =1185
                    Height =300
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6345
                    Top =60
                    Width =1185
                    Height =300
                    TabIndex =3
                    Name ="Text72"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6345
                    LayoutCachedTop =60
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7545
                    Top =60
                    Width =1185
                    Height =300
                    TabIndex =4
                    Name ="Text73"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =60
                    LayoutCachedWidth =8730
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8715
                    Top =60
                    Width =1065
                    Height =300
                    TabIndex =5
                    Name ="Text74"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8715
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Left =735
                    Top =60
                    Width =2100
                    Height =300
                    TabIndex =6
                    Name ="txtGroupLevelDesc"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            TextFontFamily =34
                            Left =120
                            Top =60
                            Width =285
                            Height =300
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
                    Left =495
                    Top =60
                    Width =180
                    Height =300
                    TabIndex =7
                    Name ="txtGroupLevel"
                    ControlSource ="GroupLvl"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Top =60
                    Width =1065
                    Height =300
                    TabIndex =8
                    Name ="Text85"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =60
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =360
                End
            End
        End
        Begin PageFooter
            Height =690
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
                    Width =4920
                    Height =330
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
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
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
                    Width =4920
                    Height =330
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
                    Width =10560
                    Height =330
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
                    Width =10560
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
            CanGrow = NotDefault
            Height =1185
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=0-Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=0-Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D1_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=0-Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D2_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8445
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=0-Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D3_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Width =1065
                    Height =330
                    FontWeight =700
                    TabIndex =4
                    Name ="D4 Grand Total Sum"
                    ControlSource ="=0-Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D4_Grand_Total_Sum"

                End
                Begin Label
                    TextFontFamily =34
                    Width =1620
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Total"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =2340
                    Width =1980
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    FontName ="Arial Narrow"
                End
                Begin Line
                    BorderWidth =3
                    Top =360
                    Width =10680
                    BorderColor =12632256
                    Name ="Line78"
                End
                Begin Subform
                    Top =450
                    Width =10681
                    Height =735
                    TabIndex =5
                    Name ="qryOLCCDepositRecap"
                    SourceObject ="Report.qryEOS040OLCCDepositRecap"

                    LayoutCachedTop =450
                    LayoutCachedWidth =10681
                    LayoutCachedHeight =1185
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD040SalesByDepartment.cls"
