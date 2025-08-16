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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =116
    Top =600
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x23545bdd8de3e240
    End
    RecordSource ="qryEOSalesByDepartmentDetail"
    Caption ="qryEOSalesByDepartmentDetail"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Department"
        End
        Begin BreakLevel
            ControlSource ="DepartmentDesc"
        End
        Begin BreakLevel
            ControlSource ="ID"
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
                    Top =540
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
                    Left =120
                    Top =60
                    Width =10560
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items Sold Today GP Report (Includes counter and bar sales)"
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
            End
        End
        Begin PageHeader
            Height =555
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =1080
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6660
                    Top =120
                    Width =1680
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Net Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8460
                    Top =120
                    Width =690
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9315
                    Top =120
                    Width =750
                    Height =300
                    Name ="D2_Label"
                    Caption ="Profit"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =10050
                    Top =120
                    Width =630
                    Height =300
                    Name ="D3_Label"
                    Caption ="GP %"
                    Tag ="DetachedLabel"
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
                    TextAlign =2
                    Left =3120
                    Top =120
                    Width =1200
                    Height =300
                    Name ="Label114"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =4680
                    Top =120
                    Width =1740
                    Height =300
                    Name ="Label115"
                    Caption ="Returns"
                    Tag ="DetachedLabel"
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
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =2
            Name ="GroupHeader2"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =225
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =2160
                    Width =1140
                    Height =225
                    Name ="Text82"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =60
                    Width =2040
                    Height =225
                    TabIndex =1
                    Name ="Text83"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =6585
                    Width =720
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
                    Left =7320
                    Width =900
                    Height =225
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="TotalP"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8250
                    Width =900
                    Height =225
                    TabIndex =4
                    Name ="TotalC"
                    ControlSource ="TotalC"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9165
                    Width =900
                    Height =225
                    TabIndex =5
                    Name ="TotalGP"
                    ControlSource ="TotalGP"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9840
                    Width =840
                    Height =225
                    TabIndex =6
                    Name ="TotalGPP"
                    ControlSource ="=IIf([TotalP]<>0,[TotalGP]/[TotalP],0)"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9840
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =2880
                    Width =735
                    Height =225
                    TabIndex =7
                    Name ="TotalSalesQ"
                    ControlSource ="TotalSalesQ"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3660
                    Width =900
                    Height =225
                    TabIndex =8
                    Name ="TotalSalesD"
                    ControlSource ="TotalSalesD"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4620
                    Width =735
                    Height =225
                    TabIndex =9
                    Name ="TotalReturnsQ"
                    ControlSource ="TotalReturnsQ"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5400
                    Width =900
                    Height =225
                    TabIndex =10
                    Name ="TotalReturnsD"
                    ControlSource ="TotalReturnsD"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            BreakLevel =2
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    Left =60
                    Top =120
                    Width =420
                    FontWeight =700
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =540
                    Top =120
                    Width =2220
                    ColumnWidth =1800
                    FontWeight =700
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =6600
                    Top =120
                    Width =720
                    FontWeight =700
                    TabIndex =2
                    Name ="TQ"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7320
                    Top =120
                    Width =900
                    FontWeight =700
                    TabIndex =3
                    Name ="TP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8250
                    Top =120
                    Width =900
                    FontWeight =700
                    TabIndex =4
                    Name ="TC"
                    ControlSource ="=Sum([TotalC])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\015\012\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9165
                    Top =120
                    Width =900
                    FontWeight =700
                    TabIndex =5
                    Name ="TGP"
                    ControlSource ="=Sum([TotalGP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9840
                    Top =120
                    Width =840
                    FontWeight =700
                    TabIndex =6
                    Name ="TGPPercent"
                    ControlSource ="=(Sum([TotalGP])/Sum([TotalP]))"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =120
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4620
                    Top =120
                    Width =720
                    FontWeight =700
                    TabIndex =7
                    Name ="Text96"
                    ControlSource ="=Sum([TotalReturnsQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5400
                    Top =120
                    Width =900
                    FontWeight =700
                    TabIndex =8
                    Name ="Text97"
                    ControlSource ="=Sum([TotalReturnsD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =2880
                    Top =120
                    Width =720
                    FontWeight =700
                    TabIndex =9
                    Name ="Text98"
                    ControlSource ="=Sum([TotalSalesQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3660
                    Top =120
                    Width =900
                    FontWeight =700
                    TabIndex =10
                    Name ="Text99"
                    ControlSource ="=Sum([TotalSalesD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =120
                    Top =60
                    Width =10560
                    BorderColor =12632256
                    Name ="Line113"
                End
                Begin Line
                    Left =240
                    Top =360
                    Width =10440
                    BorderColor =12632256
                    Name ="Line112"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =60
                    Top =180
                    Width =300
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="*"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    Left =6405
                    Top =180
                    Width =900
                    Height =270
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TotalQ"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7320
                    Top =180
                    Width =900
                    Height =270
                    FontWeight =700
                    TabIndex =1
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TotalP"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8250
                    Top =180
                    Width =900
                    Height =270
                    TabIndex =2
                    Name ="Sum Of D1"
                    ControlSource ="=Sum([TotalC])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D1"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9165
                    Top =180
                    Width =900
                    Height =270
                    TabIndex =3
                    Name ="Sum Of D2"
                    ControlSource ="=Sum([TotalGP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_D2"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9840
                    Top =180
                    Width =840
                    Height =270
                    TabIndex =4
                    Name ="GGP"
                    ControlSource ="=(Sum([TotalGP])/Sum([TotalP]))"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =180
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =450
                End
                Begin Line
                    Left =120
                    Top =60
                    Width =10560
                    BorderColor =12632256
                    Name ="Line58"
                End
                Begin TextBox
                    Left =420
                    Top =180
                    Width =2460
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4620
                    Top =180
                    Width =720
                    FontWeight =700
                    TabIndex =6
                    Name ="Text100"
                    ControlSource ="=Sum([TotalReturnsQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5400
                    Top =180
                    Width =900
                    FontWeight =700
                    TabIndex =7
                    Name ="Text101"
                    ControlSource ="=Sum([TotalReturnsD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =2880
                    Top =180
                    Width =720
                    FontWeight =700
                    TabIndex =8
                    Name ="Text102"
                    ControlSource ="=Sum([TotalSalesQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3660
                    Top =180
                    Width =900
                    FontWeight =700
                    TabIndex =9
                    Name ="Text103"
                    ControlSource ="=Sum([TotalSalesD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    Left =6405
                    Top =60
                    Width =900
                    Height =300
                    Name ="Text69"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    Left =7320
                    Top =60
                    Width =900
                    Height =300
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8250
                    Top =60
                    Width =900
                    Height =300
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="=Sum([TotalC])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9165
                    Top =60
                    Width =900
                    Height =300
                    TabIndex =3
                    Name ="Text72"
                    ControlSource ="=Sum([TotalGP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9840
                    Top =60
                    Width =840
                    Height =300
                    TabIndex =4
                    Name ="Text73"
                    ControlSource ="=(Sum([TotalGP])/Sum([TotalP]))"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Left =735
                    Top =60
                    Width =2100
                    Height =300
                    FontSize =12
                    TabIndex =5
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
                    FontSize =12
                    TabIndex =6
                    Name ="txtGroupLevel"
                    ControlSource ="GroupLvl"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4620
                    Top =60
                    Width =720
                    FontWeight =700
                    TabIndex =7
                    Name ="Text104"
                    ControlSource ="=Sum([TotalReturnsQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5400
                    Top =60
                    Width =900
                    FontWeight =700
                    TabIndex =8
                    Name ="Text105"
                    ControlSource ="=Sum([TotalReturnsD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =2880
                    Top =60
                    Width =720
                    FontWeight =700
                    TabIndex =9
                    Name ="Text106"
                    ControlSource ="=Sum([TotalSalesQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3660
                    Top =60
                    Width =900
                    FontWeight =700
                    TabIndex =10
                    Name ="Text107"
                    ControlSource ="=Sum([TotalSalesD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

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
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7320
                    Width =900
                    Height =300
                    FontWeight =700
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8250
                    Width =900
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([TotalC])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D1_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9165
                    Width =900
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([TotalGP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D2_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    TextAlign =3
                    Left =9840
                    Width =840
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=(Sum([TotalGP])/Sum([TotalP]))"
                    Format ="Percent"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D3_Grand_Total_Sum"

                    LayoutCachedLeft =9840
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =300
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
                    Left =960
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
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    Left =6405
                    Width =900
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="Text91"
                    ControlSource ="=Sum([TotalQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4620
                    Top =60
                    Width =720
                    FontWeight =700
                    TabIndex =5
                    Name ="Text108"
                    ControlSource ="=Sum([TotalReturnsQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5400
                    Top =60
                    Width =900
                    FontWeight =700
                    TabIndex =6
                    Name ="Text109"
                    ControlSource ="=Sum([TotalReturnsD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =2880
                    Top =60
                    Width =720
                    FontWeight =700
                    TabIndex =7
                    Name ="Text110"
                    ControlSource ="=Sum([TotalSalesQ])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3660
                    Top =60
                    Width =900
                    FontWeight =700
                    TabIndex =8
                    Name ="Text111"
                    ControlSource ="=Sum([TotalSalesD])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD120RItemsSoldToday.cls"
