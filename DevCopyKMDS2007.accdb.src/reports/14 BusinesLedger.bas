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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =24
    RecSrcDt = Begin
        0x157e34dd0809e340
    End
    RecordSource ="qryBusinessLedger"
    Caption ="qryBusinessLedger"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =840
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Width =10620
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Business Ledger"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =15
                    Top =540
                    Width =10621
                    Height =300
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =300
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Width =690
                    Height =225
                    Name ="Label1"
                    Caption ="BusDate"
                End
                Begin Label
                    TextAlign =3
                    Left =1950
                    Width =1050
                    Height =225
                    Name ="Label3"
                    Caption ="SalesCounter"
                End
                Begin Label
                    TextAlign =3
                    Left =3450
                    Width =1050
                    Height =225
                    Name ="Label5"
                    Caption ="SalesClubs"
                End
                Begin Label
                    TextAlign =3
                    Left =4950
                    Width =1050
                    Height =225
                    Name ="Label7"
                    Caption ="Purchases"
                End
                Begin Label
                    TextAlign =3
                    Left =6450
                    Width =1050
                    Height =225
                    Name ="Label9"
                    Caption ="Transfer IN"
                End
                Begin Label
                    TextAlign =3
                    Left =9480
                    Width =1050
                    Height =225
                    Name ="Label11"
                    Caption ="Adjust"
                End
                Begin Label
                    TextAlign =3
                    Left =7935
                    Width =1065
                    Height =225
                    Name ="Label21"
                    Caption ="Transfer OUT"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Name ="BusDate"
                    ControlSource ="BusDate"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1560
                    TabIndex =1
                    Name ="SalesCounter"
                    ControlSource ="SalesCounter"
                    Format ="#,##0.00;-#,##0.00[Red];0.00;0.00"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3060
                    TabIndex =2
                    Name ="SalesClubs"
                    ControlSource ="SalesClubs"
                    Format ="#,##0.00;-#,##0.00[Red];0.00;0.00"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4560
                    TabIndex =3
                    Name ="Purchases"
                    ControlSource ="Purchases"
                    Format ="#,##0.00;-#,##0.00[Red];0.00;0.00"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6060
                    TabIndex =4
                    Name ="TransfersIn"
                    ControlSource ="TransfersIn"
                    Format ="#,##0.00;-#,##0.00[Red];0.00;0.00"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9120
                    TabIndex =5
                    Name ="Adjust"
                    ControlSource ="Adjust"
                    Format ="#,##0.00;-#,##0.00[Red];0.00;0.00"

                End
                Begin Line
                    OldBorderStyle =0
                    Left =60
                    Top =240
                    Width =9000
                    Name ="line5"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7560
                    TabIndex =6
                    Name ="TransfersOut"
                    ControlSource ="TransfersOut"
                    Format ="#,##0.00;-#,##0.00[Red];0.00;0.00"

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
                    BackStyle =0
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
                    BackStyle =0
                    Left =5640
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
                    BackStyle =0
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
                    BackStyle =0
                    Left =5640
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
                    BackStyle =0
                    Top =300
                    Width =10560
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
                    BackStyle =0
                    Left =45
                    Width =10560
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
            Height =420
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1560
                    Top =180
                    FontWeight =700
                    Name ="Text13"
                    ControlSource ="=Sum([SalesCounter])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3060
                    Top =180
                    FontWeight =700
                    TabIndex =1
                    Name ="Text14"
                    ControlSource ="=Sum([SalesClubs])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4560
                    Top =180
                    FontWeight =700
                    TabIndex =2
                    Name ="Text15"
                    ControlSource ="=Sum([Purchases])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6060
                    Top =180
                    FontWeight =700
                    TabIndex =3
                    Name ="Text16"
                    ControlSource ="=Sum([TransfersIn])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9120
                    Top =180
                    FontWeight =700
                    TabIndex =4
                    Name ="Text17"
                    ControlSource ="=Sum([Adjust])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin Line
                    Left =120
                    Top =60
                    Width =10500
                    Name ="Line18"
                End
                Begin Line
                    Left =120
                    Width =10500
                    Name ="Line19"
                End
                Begin Label
                    TextFontFamily =34
                    Left =180
                    Top =180
                    Width =690
                    Height =225
                    FontWeight =700
                    Name ="Label20"
                    Caption ="Total"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7560
                    Top =180
                    FontWeight =700
                    TabIndex =5
                    Name ="Text23"
                    ControlSource ="=Sum([TransfersOut])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
            End
        End
    End
End
CodeBehindForm
' See "14 BusinesLedger.cls"
