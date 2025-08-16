Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =69
    Top =600
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xcbe6ab71d6a2e240
    End
    RecordSource ="qry EOM110SalePricesReport"
    Caption ="qry EOM110SalePricesReport"
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
                    Left =2280
                    Top =60
                    Width =6225
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOM Sale Prices"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10740
                    Height =30
                    BorderColor =12632256
                    Name ="Line50"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =10740
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10740
                    Height =30
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =720
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =900
                    Top =180
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption =" Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3720
                    Top =180
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Old"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4725
                    Top =180
                    Width =1035
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sale Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6360
                    Top =180
                    Width =900
                    Height =300
                    Name ="D1_Label"
                    Caption =" Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7380
                    Top =180
                    Width =1020
                    Height =300
                    Name ="D2_Label"
                    Caption ="Old Value"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8535
                    Top =180
                    Width =1095
                    Height =300
                    Name ="D3_Label"
                    Caption ="New Value"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =10200
                    Top =180
                    Width =495
                    Height =300
                    Name ="D4_Label"
                    Caption =" Diff"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1020
                    Top =90
                    Width =9660
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1020
                    Top =60
                    Width =9660
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1020
                    Top =660
                    Width =9720
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1020
                    Top =600
                    Width =9720
                    Name ="Line68"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =420
                    Width =1260
                    Height =270
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =420
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Left =1860
                    Width =1740
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1860
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    Left =3360
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="Old"
                    ControlSource ="Old"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Width =1185
                    Height =270
                    FontWeight =700
                    TabIndex =3
                    Name ="New"
                    ControlSource ="Sale"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4545
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =6060
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6060
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Width =1185
                    Height =270
                    TabIndex =5
                    Name ="OldValue"
                    ControlSource ="OldValue"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7245
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Width =1185
                    Height =270
                    TabIndex =6
                    Name ="NewValue"
                    ControlSource ="NewValue"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8430
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Width =1125
                    Height =300
                    TabIndex =7
                    Name ="Markup"
                    ControlSource ="Markup"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9615
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =300
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
                    Width =4980
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
                    Width =10740
                    Height =30
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10740
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
                    Width =4980
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
                    Width =10620
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
                    Width =10620
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
            Height =570
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([Onhand])"
                    Format ="Standard"
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
                    TabIndex =1
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([OldValue])"
                    Format ="Standard"
                    EventProcPrefix ="D2_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=Sum([NewValue])"
                    Format ="Standard"
                    EventProcPrefix ="D3_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Width =1125
                    Height =330
                    FontWeight =700
                    TabIndex =3
                    Name ="D4 Grand Total Sum"
                    ControlSource ="=Sum([Markup])"
                    Format ="Standard"
                    EventProcPrefix ="D4_Grand_Total_Sum"

                End
                Begin Label
                    TextAlign =3
                    Left =1980
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption =" Total"
                End
                Begin Label
                    Visible = NotDefault
                    Top =300
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
' See "qryEOM110SalePricesReport.cls"
