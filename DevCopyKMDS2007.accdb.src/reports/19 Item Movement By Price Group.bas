Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =61
    Left =1170
    Top =2040
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xb9f5fcf5352ae340
    End
    RecordSource ="qryItemMovementByPriceGroupReport"
    Caption ="Item Movement By Price Group"
    DatasheetFontName ="Trebuchet MS"
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
            ControlSource ="PriceGroupDesc"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="ItemID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1380
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10560
                    Height =510
                    FontSize =20
                    Name ="Label43"
                    Caption ="Item Movement By Price Group"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line46"
                End
                Begin TextBox
                    TextAlign =2
                    Top =600
                    Width =10621
                    Height =300
                    ColumnOrder =0
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Top =900
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text59"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboPriceGroupIDBegin.Column(1) & \" a"
                        "nd \" & Forms!PrintInventoryReport!cboPriceGroupIDEnd.Column(1)"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1125
                    Height =300
                    Name ="PriceGroupID_Label"
                    Caption ="PriceGroup"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1440
                    Top =60
                    Width =1140
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5280
                    Top =60
                    Width =1080
                    Height =300
                    Name ="Purchases_Label"
                    Caption ="Purchases"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6420
                    Top =60
                    Width =840
                    Height =300
                    Name ="Counter_Label"
                    Caption ="Counter"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7320
                    Top =60
                    Width =840
                    Height =300
                    Name ="Clubs_Label"
                    Caption ="Clubs"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8220
                    Top =60
                    Width =1020
                    Height =300
                    Name ="Transfers_Label"
                    Caption ="Transfers"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9300
                    Top =60
                    Width =1200
                    Height =300
                    Name ="Adjustments_Label"
                    Caption ="Adjustments"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line47"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10800
                    Name ="Line48"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =390
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Width =2160
                    Height =270
                    Name ="PriceGroupDesc"
                    ControlSource ="PriceGroupDesc"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =2160
                    Width =600
                    Height =270
                    ColumnWidth =810
                    Name ="ID"
                    ControlSource ="ItemID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2940
                    Width =540
                    Height =270
                    ColumnWidth =2190
                    TabIndex =1
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3660
                    Width =540
                    Height =270
                    TabIndex =2
                    Name ="Purchases"
                    ControlSource ="Purchases"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5580
                    Width =540
                    Height =270
                    TabIndex =3
                    Name ="Counter"
                    ControlSource ="Counter"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6180
                    Width =540
                    Height =270
                    TabIndex =4
                    Name ="Clubs"
                    ControlSource ="Clubs"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6780
                    Width =540
                    Height =270
                    TabIndex =5
                    Name ="Transfers"
                    ControlSource ="Transfers"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7380
                    Width =600
                    Height =270
                    TabIndex =6
                    Name ="Adjustments"
                    ControlSource ="Adjustments"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8040
                    Width =1140
                    Height =270
                    ColumnWidth =780
                    TabIndex =7
                    Name ="Price"
                    ControlSource ="Price"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9240
                    Width =600
                    Height =270
                    ColumnWidth =1230
                    TabIndex =8
                    Name ="InputCode"
                    ControlSource ="InputCode"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9900
                    Width =840
                    Height =270
                    TabIndex =9
                    Name ="Emp"
                    ControlSource ="Emp"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Width =1140
                    Height =270
                    ColumnWidth =2190
                    TabIndex =10
                    Name ="Description"
                    ControlSource ="Description"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =270
            BreakLevel =1
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =5280
                    Width =1080
                    Height =225
                    Name ="Text50"
                    ControlSource ="=Sum([Purchases])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6420
                    Width =840
                    Height =225
                    TabIndex =1
                    Name ="Text51"
                    ControlSource ="=Sum([Counter])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7320
                    Width =840
                    Height =225
                    TabIndex =2
                    Name ="Text52"
                    ControlSource ="=Sum([Clubs])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8220
                    Width =1020
                    Height =225
                    TabIndex =3
                    Name ="Text53"
                    ControlSource ="=Sum([Transfers])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9300
                    Width =1200
                    Height =225
                    TabIndex =4
                    Name ="Text54"
                    ControlSource ="=Sum([Adjustments])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2220
                    Width =2820
                    Height =225
                    TabIndex =5
                    Name ="Text55"
                    ControlSource ="Description"

                    LayoutCachedLeft =2220
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Width =1260
                    Height =225
                    TabIndex =6
                    Name ="Text60"
                    ControlSource ="ItemID"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5340
                    Top =150
                    Width =1080
                    Height =270
                    FontWeight =700
                    Name ="Sum Of Purchases"
                    ControlSource ="=Sum([Purchases])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_Purchases"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =6480
                    Top =150
                    Width =840
                    Height =270
                    FontWeight =700
                    TabIndex =1
                    Name ="Sum Of Counter"
                    ControlSource ="=Sum([Counter])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_Counter"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =7380
                    Top =150
                    Width =840
                    Height =270
                    FontWeight =700
                    TabIndex =2
                    Name ="Sum Of Clubs"
                    ControlSource ="=Sum([Clubs])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_Clubs"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =8280
                    Top =150
                    Width =1020
                    Height =270
                    FontWeight =700
                    TabIndex =3
                    Name ="Sum Of Transfers"
                    ControlSource ="=Sum([Transfers])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_Transfers"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =9360
                    Top =150
                    Width =1200
                    Height =270
                    FontWeight =700
                    TabIndex =4
                    Name ="Sum Of Adjustments"
                    ControlSource ="=Sum([Adjustments])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_Adjustments"

                End
                Begin Line
                    Left =900
                    Top =480
                    Width =9840
                    Name ="Line56"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =3180
                    Top =120
                    Width =2160
                    Height =270
                    FontWeight =700
                    TabIndex =5
                    Name ="Text57"
                    ControlSource ="PriceGroupDesc"

                End
                Begin Line
                    Left =5400
                    Top =120
                    Width =5160
                    Name ="Line58"
                End
            End
        End
        Begin PageFooter
            Height =615
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
                    Width =4320
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
            Height =360
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5280
                    Width =1080
                    Height =300
                    FontSize =10
                    Name ="Purchases Grand Total Sum"
                    ControlSource ="=Sum([Purchases])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    FontName ="Times New Roman"
                    EventProcPrefix ="Purchases_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6420
                    Width =840
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Counter Grand Total Sum"
                    ControlSource ="=Sum([Counter])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    FontName ="Times New Roman"
                    EventProcPrefix ="Counter_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label32"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7320
                    Width =840
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Clubs Grand Total Sum"
                    ControlSource ="=Sum([Clubs])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    FontName ="Times New Roman"
                    EventProcPrefix ="Clubs_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label35"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8220
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="Transfers Grand Total Sum"
                    ControlSource ="=Sum([Transfers])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    FontName ="Times New Roman"
                    EventProcPrefix ="Transfers_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label38"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9300
                    Width =1200
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Adjustments Grand Total Sum"
                    ControlSource ="=Sum([Adjustments])"
                    Format ="#,###;-#,###[Red];\"\";\"Null\""
                    FontName ="Times New Roman"
                    EventProcPrefix ="Adjustments_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label41"
                    Caption ="Grand Total"
                End
            End
        End
    End
End
CodeBehindForm
' See "19 Item Movement By Price Group.cls"
