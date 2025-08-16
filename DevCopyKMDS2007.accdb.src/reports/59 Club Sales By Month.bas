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
    ItemSuffix =160
    Left =1050
    Top =1860
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xb0b7f592ee54e340
    End
    RecordSource ="qryClubInvoiceHistoryByInvoiceDetailYYYYMM"
    Caption ="qryClubInvoiceHistoryByInvoiceDetail"
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
            GroupFooter = NotDefault
            ControlSource ="ClubName"
        End
        Begin BreakLevel
            ControlSource ="YYYYMM"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =10695
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Club Invoices"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10785
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10785
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10621
                    Height =480
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin PageHeader
            Height =825
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =30
                    Width =10740
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Width =10740
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =780
                    Width =10740
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =810
                    Width =10740
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =120
                    Width =2340
                    Height =300
                    Name ="Label58"
                    Caption ="*****    Liquor   *****"
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =420
                    Width =2340
                    Height =300
                    Name ="Label59"
                    Caption ="Retail    Disct      Net"
                End
                Begin Label
                    TextAlign =3
                    Left =120
                    Top =420
                    Width =570
                    Height =300
                    Name ="Label61"
                    Caption ="Club"
                End
                Begin Label
                    TextAlign =3
                    Left =6285
                    Top =120
                    Width =2115
                    Height =300
                    Name ="Label90"
                    Caption ="*** Non Liquor   ***"
                End
                Begin Label
                    TextAlign =3
                    Left =6240
                    Top =420
                    Width =2160
                    Height =300
                    Name ="Label91"
                    Caption ="Retail    Disct      Net"
                End
                Begin Label
                    TextAlign =3
                    Left =8400
                    Top =120
                    Width =2295
                    Height =300
                    Name ="Label101"
                    Caption ="***       Total    ***"
                End
                Begin Label
                    TextAlign =3
                    Left =8415
                    Top =420
                    Width =2280
                    Height =300
                    Name ="Label102"
                    Caption ="Retail    Disct      Net"
                End
                Begin Rectangle
                    Left =3780
                    Top =120
                    Width =2400
                    Height =600
                    Name ="Box118"
                End
                Begin Rectangle
                    Left =6300
                    Top =120
                    Width =2160
                    Height =600
                    Name ="Box119"
                End
                Begin Rectangle
                    Left =8640
                    Top =120
                    Width =2100
                    Height =600
                    Name ="Box120"
                End
                Begin Label
                    TextAlign =3
                    Left =2115
                    Top =420
                    Width =1335
                    Height =300
                    Name ="Label146"
                    Caption ="Year - Month"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Width =2520
                    Height =225
                    Name ="ClubName"
                    ControlSource ="ClubName"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2880
                    Width =720
                    Height =225
                    ColumnWidth =2400
                    TabIndex =1
                    Name ="SetupDate"
                    ControlSource ="YYYYMM"
                    Format ="mm/dd/yy"

                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10380
                    Name ="Line5"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Width =540
                    Height =255
                    TabIndex =2
                    Name ="DisDollars"
                    ControlSource ="LiquorDiscount"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4080
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="InvDollars"
                    ControlSource ="GrossLiquor"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5460
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="NetDollars"
                    ControlSource ="=[GrossLiquor]+[LiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Width =540
                    Height =255
                    TabIndex =5
                    Name ="Text92"
                    ControlSource ="NonLiquorDiscount"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text93"
                    ControlSource ="GrossNonLiquor"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7800
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="Text94"
                    ControlSource ="=[GrossNonLiquor]+[NonLiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Width =540
                    Height =255
                    TabIndex =8
                    Name ="Text103"
                    ControlSource ="=[NonLiquorDiscount]+[LiquorDiscount]"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Width =720
                    Height =225
                    TabIndex =9
                    Name ="Text104"
                    ControlSource ="=[GrossNonLiquor]+[GrossLiquor]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10020
                    Width =720
                    Height =225
                    TabIndex =10
                    Name ="Text105"
                    ControlSource ="=[GrossNonLiquor]+[NonLiquorDiscount]+[GrossLiquor]+[LiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =405
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Top =180
                    Width =2520
                    Height =225
                    Name ="Text147"
                    ControlSource ="ClubName"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =180
                    Width =540
                    Height =225
                    TabIndex =1
                    Name ="Text149"
                    ControlSource ="=Sum([LiquorDiscount])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4080
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Text150"
                    ControlSource ="=Sum([GrossLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5460
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Text151"
                    ControlSource ="=Sum([GrossLiquor]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Top =180
                    Width =540
                    Height =225
                    TabIndex =4
                    Name ="Text152"
                    ControlSource ="=Sum([NonLiquorDiscount])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Text153"
                    ControlSource ="=Sum([GrossNonLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7800
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text154"
                    ControlSource ="=Sum([GrossNonLiquor]+[NonLiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Top =180
                    Width =540
                    Height =225
                    TabIndex =7
                    Name ="Text155"
                    ControlSource ="=Sum([NonLiquorDiscount]+[LiquorDiscount])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =8
                    Name ="Text156"
                    ControlSource ="=Sum([GrossNonLiquor]+[GrossLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10020
                    Top =180
                    Width =720
                    Height =225
                    TabIndex =9
                    Name ="Text157"
                    ControlSource ="=Sum([GrossNonLiquor]+[NonLiquorDiscount]+[GrossLiquor]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =3540
                    Top =60
                    Width =7200
                    Name ="Line158"
                End
                Begin Line
                    Left =3540
                    Top =120
                    Width =7200
                    Name ="Line159"
                End
            End
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
                    ControlSource ="=[ClubID]"
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
            Height =1260
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Label
                    Left =1800
                    Top =360
                    Width =1080
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin Line
                    Left =3540
                    Top =180
                    Width =7200
                    Name ="Line70"
                End
                Begin Label
                    Visible = NotDefault
                    Top =360
                    Width =2100
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
                Begin Line
                    Left =3540
                    Top =240
                    Width =7200
                    Name ="Line81"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3840
                    Top =660
                    Width =1200
                    Name ="Text87"
                    ControlSource ="=Sum([LiquorDiscount])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3840
                    Top =360
                    Width =1200
                    TabIndex =1
                    Name ="Text88"
                    ControlSource ="=Sum([GrossLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3840
                    Top =1020
                    Width =1200
                    TabIndex =2
                    Name ="Text89"
                    ControlSource ="=Sum([GrossLiquor]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Top =660
                    Width =1200
                    TabIndex =3
                    Name ="Text98"
                    ControlSource ="=Sum([NonLiquorDiscount])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Top =360
                    Width =1200
                    Height =225
                    TabIndex =4
                    Name ="Text99"
                    ControlSource ="=Sum([GrossNonLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Top =1035
                    Width =1200
                    Height =225
                    TabIndex =5
                    Name ="Text100"
                    ControlSource ="=Sum([GrossNonLiquor]+[NonLiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =660
                    Width =1200
                    TabIndex =6
                    Name ="Text115"
                    ControlSource ="=Sum([NonLiquorDiscount]+[LiquorDiscount])"
                    Format ="#,##0.00;#,##0.00[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =360
                    Width =1200
                    Height =225
                    TabIndex =7
                    Name ="Text116"
                    ControlSource ="=Sum([GrossNonLiquor]+[GrossLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =1020
                    Width =1200
                    TabIndex =8
                    Name ="Text117"
                    ControlSource ="=Sum([GrossNonLiquor]+[NonLiquorDiscount]+[GrossLiquor]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1980
                    Top =960
                    Width =600
                    TabIndex =9
                    Name ="txtCount"
                    Format ="General Number"

                End
                Begin Label
                    TextAlign =3
                    Left =2760
                    Top =360
                    Width =1035
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label125"
                    Caption ="Gross"
                End
                Begin Label
                    TextAlign =3
                    Left =2760
                    Top =660
                    Width =1035
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label126"
                    Caption ="Discount"
                End
                Begin Label
                    TextAlign =3
                    Left =2760
                    Top =990
                    Width =1035
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label127"
                    Caption ="Net"
                End
                Begin Line
                    Left =3540
                    Top =960
                    Width =7200
                    Name ="Line128"
                End
            End
        End
    End
End
CodeBehindForm
' See "59 Club Sales By Month.cls"
