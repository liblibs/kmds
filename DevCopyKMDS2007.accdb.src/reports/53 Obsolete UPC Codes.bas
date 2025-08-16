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
    Width =10260
    DatasheetFontHeight =10
    ItemSuffix =53
    Left =570
    Top =1290
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x6a4a9ed39142e340
    End
    RecordSource ="qryUPCItemsObsolete"
    Caption ="53 Obsolete UPC Codes"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            ControlSource ="ItemID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1260
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
                    Caption ="Obsolete UPC Codes "
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =12632256
                    Name ="Line41"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =0
                    Left =60
                    Top =690
                    Width =7200
                    Height =510
                    FontSize =14
                    Name ="Label52"
                    Caption ="UPC Codes older than 60 days and not scanned in 500 days"
                End
            End
        End
        Begin PageHeader
            Height =435
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
                    Left =3060
                    Top =60
                    Width =825
                    Height =300
                    Name ="ID_Label"
                    Caption ="Item ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6300
                    Top =60
                    Width =1920
                    Height =300
                    Name ="SetupDate_Label"
                    Caption ="Obsolete UPC Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8340
                    Top =60
                    Width =1110
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Setup Date"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =9420
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =9420
                    Name ="Line43"
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
                    ControlSource ="SalesDepartmentID"
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
                    ControlSource ="SalesDepartment.Description"
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
                    Left =3780
                    Width =2820
                    Height =225
                    ColumnWidth =3165
                    Name ="Description"
                    ControlSource ="Item.Description"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6780
                    Width =1320
                    FontWeight =700
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8280
                    Width =1980
                    ColumnWidth =2100
                    TabIndex =2
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2460
                    Width =1200
                    Height =225
                    TabIndex =3
                    Name ="PONumber"
                    ControlSource ="ItemID"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
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
            End
        End
        Begin PageFooter
            Height =300
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
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
                    ControlSource ="=Count(*)"
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
            End
        End
    End
End
CodeBehindForm
' See "53 Obsolete UPC Codes.cls"
