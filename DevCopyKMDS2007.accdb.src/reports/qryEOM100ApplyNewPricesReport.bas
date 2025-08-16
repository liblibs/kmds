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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =74
    Left =1140
    Top =1980
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xc34e4f639039e440
    End
    RecordSource ="qry EOM100ApplyNewPricesReport"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1215
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
                    Left =2295
                    Top =60
                    Width =6195
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOM Price Changes"
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
                    Top =1200
                    Width =10800
                    BorderColor =12632256
                    Name ="Line73"
                End
            End
        End
        Begin PageHeader
            Height =720
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =180
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption =" Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3600
                    Top =180
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Old"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3600
                    LayoutCachedTop =180
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =3
                    Left =4575
                    Top =180
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="New"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4575
                    LayoutCachedTop =180
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =3
                    Left =5400
                    Top =180
                    Width =900
                    Height =300
                    Name ="D1_Label"
                    Caption =" Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =180
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =3
                    Left =6435
                    Top =180
                    Width =1020
                    Height =300
                    Name ="D2_Label"
                    Caption ="Old Value"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6435
                    LayoutCachedTop =180
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =3
                    Left =7545
                    Top =180
                    Width =1095
                    Height =300
                    Name ="D3_Label"
                    Caption ="New Value"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7545
                    LayoutCachedTop =180
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =3
                    Left =8625
                    Top =180
                    Width =930
                    Height =300
                    Name ="D4_Label"
                    Caption ="Markups"
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
                Begin Label
                    TextAlign =3
                    Left =9420
                    Top =180
                    Width =1200
                    Height =300
                    Name ="Label69"
                    Caption =" Downs"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Height =225
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedWidth =1440
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =3660
                    Width =765
                    Height =225
                    TabIndex =2
                    Name ="Old"
                    ControlSource ="Old"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3660
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4560
                    Width =765
                    Height =225
                    FontWeight =700
                    TabIndex =3
                    Name ="New"
                    ControlSource ="New"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4560
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =5520
                    Width =765
                    Height =225
                    TabIndex =4
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5520
                    LayoutCachedWidth =6285
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6300
                    Width =1185
                    Height =225
                    TabIndex =5
                    Name ="OldValue"
                    ControlSource ="OldValue"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =7485
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7440
                    Width =1185
                    Height =225
                    TabIndex =6
                    Name ="NewValue"
                    ControlSource ="NewValue"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7440
                    LayoutCachedWidth =8625
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8475
                    Width =1125
                    Height =225
                    TabIndex =7
                    Name ="Markup"
                    ControlSource ="Markup"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9540
                    Width =1125
                    Height =225
                    TabIndex =8
                    Name ="MarkDown"
                    ControlSource ="MarkDown"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Visible = NotDefault
                    BorderWidth =1
                    Top =240
                    Width =10680
                    Name ="Line5"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Left =1500
                    Width =2580
                    Height =225
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1500
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =225
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
                    Left =4920
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
                    Left =6105
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
                    Left =7290
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
                    Left =8475
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
                    Left =1140
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
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9540
                    Width =1125
                    Height =330
                    FontWeight =700
                    TabIndex =4
                    Name ="Text71"
                    ControlSource ="=Sum([MarkDown])"
                    Format ="Standard"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOM100ApplyNewPricesReport.cls"
