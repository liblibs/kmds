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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =57
    Left =960
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x2520124e9228e440
    End
    RecordSource ="qryPurchasesByVendorByDepartmentByItem"
    Caption ="52 Purchases By Vendor By Item"
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
            ControlSource ="Department"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="PriceGroupDescription"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="Description"
        End
        Begin BreakLevel
            ControlSource ="Description"
        End
        Begin BreakLevel
            ControlSource ="SetupDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =0
                    Left =60
                    Top =60
                    Width =5565
                    Height =510
                    FontSize =20
                    Name ="Label38"
                    Caption ="52 Purchases By Vendor By Item"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =12632256
                    Name ="Line41"
                End
            End
        End
        Begin PageHeader
            Height =780
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =0
                    Left =60
                    Top =60
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    Left =1440
                    Top =60
                    Width =1260
                    Height =300
                    Name ="PriceGroupID_Label"
                    Caption ="Price Group"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3720
                    Top =60
                    Width =825
                    Height =300
                    Name ="ID_Label"
                    Caption ="Item ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6960
                    Top =60
                    Width =720
                    Height =300
                    Name ="SetupDate_Label"
                    Caption ="Vendor"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7980
                    Top =60
                    Width =885
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9000
                    Top =60
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Retail"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10680
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10680
                    Name ="Line43"
                End
                Begin Label
                    TextAlign =3
                    Left =4605
                    Top =60
                    Width =1845
                    Height =300
                    Name ="Label46"
                    Caption ="PO Number - Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9900
                    Top =60
                    Width =720
                    Height =300
                    Name ="Label51"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9900
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =360
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =420
                    Height =360
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="Department"
                    ControlSource ="Department"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =540
                    Width =3480
                    Height =330
                    ColumnWidth =1635
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =390
            BreakLevel =1
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =3540
                    Height =330
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="PriceGroupID"
                    ControlSource ="PriceGroupDescription"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =2820
                    Height =225
                    ColumnWidth =3165
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2760
                    Width =1140
                    Height =225
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2760
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5100
                    Width =1215
                    Height =225
                    ColumnWidth =2100
                    TabIndex =2
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="mm/dd/yy"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5100
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7740
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="TotalQ"
                    ControlSource ="TotalQ"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8880
                    Width =840
                    Height =225
                    TabIndex =4
                    Name ="TotalP"
                    ControlSource ="TotalP"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6480
                    Width =1455
                    Height =225
                    ColumnWidth =1980
                    TabIndex =5
                    Name ="InputCode"
                    ControlSource ="InputCode"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6480
                    LayoutCachedWidth =7935
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3960
                    Width =1200
                    Height =225
                    TabIndex =6
                    Name ="PONumber"
                    ControlSource ="PONumber"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3960
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9900
                    Width =840
                    Height =225
                    TabIndex =7
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =2
            Name ="GroupFooter2"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7740
                    Top =60
                    Width =840
                    Height =300
                    Name ="Text53"
                    ControlSource ="=Sum([TotalQ])"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8700
                    Top =60
                    Width =1020
                    Height =300
                    TabIndex =1
                    Name ="Text54"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =60
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =60
                    Width =1020
                    Height =300
                    TabIndex =2
                    Name ="Text55"
                    ControlSource ="=Sum([Cost])"
                    Format ="Standard"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =7740
                    Top =60
                    Width =3000
                    Name ="Line56"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =60
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            BreakLevel =1
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =1320
                    Width =7980
                    Name ="Text28"
                    ControlSource ="=\"Summary for \" & \"'PriceGroupID' = \" & \" \" & [PriceGroupID] & \" (\" & Co"
                        "unt(*) & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =5100
                    Top =330
                    Width =2280
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Price Group TOTAL"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Top =300
                    Width =840
                    Height =300
                    TabIndex =1
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="Sum_Of_TotalQ"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8700
                    Top =300
                    Width =1020
                    Height =300
                    TabIndex =2
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"
                    EventProcPrefix ="Sum_Of_TotalP"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =300
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="Text48"
                    ControlSource ="=Sum([Cost])"
                    Format ="Standard"

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
                    IMESentenceMode =3
                    Left =60
                    Width =9240
                    Name ="Text26"
                    ControlSource ="=\"Summary for \" & \"'Department' = \" & \" \" & [Department] & \" (\" & Count("
                        "*) & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =4920
                    Top =330
                    Width =2085
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label27"
                    Caption ="Sales Department Totals"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Top =300
                    Width =840
                    Height =300
                    TabIndex =1
                    Name ="Sum Of TotalQ1"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="Sum_Of_TotalQ1"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8700
                    Top =300
                    Width =1020
                    Height =300
                    TabIndex =2
                    Name ="Sum Of TotalP1"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"
                    EventProcPrefix ="Sum_Of_TotalP1"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =300
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="Text49"
                    ControlSource ="=Sum([Cost])"
                    Format ="Standard"

                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text39"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text40"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =390
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7140
                    Top =60
                    Width =840
                    Height =330
                    FontSize =10
                    Name ="TotalQ Grand Total Sum"
                    ControlSource ="=Sum([TotalQ])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="TotalQ_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label31"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =1020
                    Height =330
                    FontSize =10
                    TabIndex =1
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([TotalP])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                End
                Begin Label
                    Left =4980
                    Top =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label35"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9720
                    Top =60
                    Width =1020
                    Height =330
                    FontSize =10
                    TabIndex =2
                    Name ="Text50"
                    ControlSource ="=Sum([Cost])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                End
            End
        End
    End
End
CodeBehindForm
' See "52 Purchases By Vendor By Item.cls"
