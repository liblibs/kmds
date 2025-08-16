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
    ItemSuffix =123
    Top =720
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xae076b547443e340
    End
    RecordSource ="qry EODClubInvoices"
    Caption ="EODClubInvoices"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
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
            End
        End
        Begin PageHeader
            Height =945
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =150
                    Width =10740
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10740
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =900
                    Width =10740
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =930
                    Width =10740
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3480
                    Top =240
                    Width =2340
                    Height =300
                    Name ="Label58"
                    Caption ="*****    Liquor   *****"
                End
                Begin Label
                    TextAlign =3
                    Left =3480
                    Top =540
                    Width =2340
                    Height =300
                    Name ="Label59"
                    Caption ="Retail    Disct         Net"
                End
                Begin Label
                    TextAlign =3
                    Left =120
                    Top =540
                    Width =570
                    Height =300
                    Name ="Label61"
                    Caption ="Club"
                End
                Begin Label
                    Left =2580
                    Top =540
                    Width =870
                    Height =315
                    Name ="Label62"
                    Caption ="Invoice"
                End
                Begin Label
                    TextAlign =3
                    Left =6060
                    Top =240
                    Width =2280
                    Height =300
                    Name ="Label90"
                    Caption ="***    Non Liquor   ***"
                End
                Begin Label
                    TextAlign =3
                    Left =6000
                    Top =540
                    Width =2340
                    Height =300
                    Name ="Label91"
                    Caption ="Retail    Disct         Net"
                End
                Begin Label
                    TextAlign =3
                    Left =8445
                    Top =240
                    Width =2295
                    Height =300
                    Name ="Label101"
                    Caption ="***       Total       ***"
                End
                Begin Label
                    TextAlign =3
                    Left =8460
                    Top =540
                    Width =2280
                    Height =300
                    Name ="Label102"
                    Caption ="Retail    Disct         Net"
                End
                Begin Rectangle
                    Left =3480
                    Top =240
                    Width =2400
                    Height =600
                    Name ="Box118"
                End
                Begin Rectangle
                    Left =6060
                    Top =240
                    Width =2340
                    Height =600
                    Name ="Box119"
                End
                Begin Rectangle
                    Left =8460
                    Top =240
                    Width =2340
                    Height =600
                    Name ="Box120"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2760
                    Width =720
                    FontSize =6
                    Name ="InvoiceNo"
                    ControlSource ="InvoiceNumber"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =30
                    Width =615
                    Height =225
                    FontSize =6
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="ClubID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =30
                    LayoutCachedWidth =645
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4320
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =2
                    Name ="DisDollars"
                    ControlSource ="LiquorDiscount"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3480
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =3
                    Name ="InvDollars"
                    ControlSource ="GrossLiquor"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\";\015\012"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3480
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5040
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =4
                    Name ="NetDollars"
                    ControlSource ="=[GrossLiquor]+[LiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5040
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =735
                    Width =2040
                    Height =225
                    FontSize =6
                    TabIndex =5
                    Name ="Text71"
                    ControlSource ="ClubName"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =735
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =6
                    Name ="Text92"
                    ControlSource ="NonLiquorDiscount"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6060
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =7
                    Name ="Text93"
                    ControlSource ="GrossNonLiquor"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7560
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =8
                    Name ="Text94"
                    ControlSource ="=[GrossNonLiquor]+[NonLiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7560
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =9
                    Name ="Text103"
                    ControlSource ="=[NonLiquorDiscount]+[LiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8400
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =10
                    Name ="Text104"
                    ControlSource ="=[GrossNonLiquor]+[GrossLiquor]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8400
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Width =720
                    Height =225
                    FontSize =6
                    TabIndex =11
                    Name ="Text105"
                    ControlSource ="=[GrossNonLiquor]+[NonLiquorDiscount]+[GrossLiquor]+[LiquorDiscount]"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Left =120
                    Top =240
                    Width =10620
                    Name ="Line5"
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
                    ControlSource ="=[Name]"
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
            Height =780
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =150
                    Top =420
                    Width =2010
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total    Invoices:"
                    FontName ="Arial Narrow"
                End
                Begin Line
                    Left =3540
                    Top =180
                    Width =7200
                    Name ="Line70"
                End
                Begin Label
                    Visible = NotDefault
                    Top =60
                    Width =3240
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
                    Left =4320
                    Top =420
                    Width =720
                    Height =270
                    Name ="Text87"
                    ControlSource ="=Sum([LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3345
                    Top =420
                    Width =870
                    Height =270
                    TabIndex =1
                    Name ="Text88"
                    ControlSource ="=Sum([GrossLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3345
                    LayoutCachedTop =420
                    LayoutCachedWidth =4215
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4905
                    Top =420
                    Width =870
                    Height =270
                    TabIndex =2
                    Name ="Text89"
                    ControlSource ="=Sum([GrossLiquor]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4905
                    LayoutCachedTop =420
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =420
                    Width =720
                    Height =270
                    TabIndex =3
                    Name ="Text98"
                    ControlSource ="=Sum([NonLiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6060
                    Top =420
                    Width =720
                    Height =270
                    TabIndex =4
                    Name ="Text99"
                    ControlSource ="=Sum([GrossNonLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7560
                    Top =420
                    Width =720
                    Height =270
                    TabIndex =5
                    Name ="Text100"
                    ControlSource ="=Sum([GrossNonLiquor]+[NonLiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =420
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9240
                    Top =420
                    Width =720
                    Height =270
                    TabIndex =6
                    Name ="Text115"
                    ControlSource ="=Sum([NonLiquorDiscount]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8265
                    Top =420
                    Width =870
                    Height =270
                    TabIndex =7
                    Name ="Text116"
                    ControlSource ="=Sum([GrossNonLiquor]+[GrossLiquor])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8265
                    LayoutCachedTop =420
                    LayoutCachedWidth =9135
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9825
                    Top =420
                    Width =870
                    Height =270
                    TabIndex =8
                    Name ="Text117"
                    ControlSource ="=Sum([GrossNonLiquor]+[NonLiquorDiscount]+[GrossLiquor]+[LiquorDiscount])"
                    Format ="#,##0.00;(#,##0.00)[Red];\"\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9825
                    LayoutCachedTop =420
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2280
                    Top =420
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    Name ="Text121"
                    ControlSource ="=Count([ClubID])"
                    FontName ="Arial Narrow"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD060ClubInvoices.cls"
