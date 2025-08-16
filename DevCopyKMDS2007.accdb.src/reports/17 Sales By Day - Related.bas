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
    ItemSuffix =134
    RecSrcDt = Begin
        0x71a365d32822e340
    End
    RecordSource ="qrySalesByDayForPC"
    Caption ="qryLiquorSalesByHour"
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
            Height =1380
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-75
                    Width =10755
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="Sales by Day for P & C"
                End
                Begin TextBox
                    TextAlign =2
                    Left =15
                    Top =540
                    Width =10621
                    Height =300
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Top =840
                    Width =10621
                    Height =300
                    TabIndex =1
                    Name ="Text134"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboSalesDepartmentIDBegin & \" \" & F"
                        "orms!PrintInventoryReport!cboSalesDepartmentIDBegin.Column(1) & \" and \" & Form"
                        "s!PrintInventoryReport!cboSalesDepartmentIDEnd & \" \" & Forms!PrintInventoryRep"
                        "ort!cboSalesDepartmentIDEnd.Column(1)"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =1395
                    Height =300
                    Name ="Date_Label"
                    Caption ="Week Begin"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =1470
                    Top =60
                    Width =1230
                    Height =300
                    Name ="SumOfPrice_Label"
                    Caption =" Week Total"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2880
                    Top =60
                    Width =720
                    Height =300
                    Name ="10 AM_Label"
                    Caption ="Sun"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl10_AM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =3900
                    Top =60
                    Width =720
                    Height =300
                    Name ="11 AM_Label"
                    Caption ="Mon"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl11_AM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =4920
                    Top =60
                    Width =720
                    Height =300
                    Name ="12 AM_Label"
                    Caption ="Tue"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl12_AM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =660
                    Height =300
                    Name ="1 PM_Label"
                    Caption ="Wed"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl1_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =6840
                    Top =60
                    Width =660
                    Height =300
                    Name ="3 PM_Label"
                    Caption ="Thu"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl3_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =7980
                    Top =60
                    Width =660
                    Height =300
                    Name ="4 PM_Label"
                    Caption ="Fri"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl4_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9360
                    Top =60
                    Width =660
                    Height =300
                    Name ="5 PM_Label"
                    Caption ="Sat"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl5_PM_Label"
                End
                Begin Line
                    Top =420
                    Width =10680
                    BorderColor =0
                    Name ="Line132"
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
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1800
                    Width =900
                    Height =225
                    Name ="SumOfPrice"
                    ControlSource ="TotalSales"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2880
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="10 AM"
                    ControlSource ="Sun"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3960
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="11 AM"
                    ControlSource ="Mon"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4980
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="12 AM"
                    ControlSource ="Tue"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="1 PM"
                    ControlSource ="Wed"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="3 PM"
                    ControlSource ="Thu"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7980
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="4 PM"
                    ControlSource ="Fri"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="5 PM"
                    ControlSource ="Sat"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl5_PM"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Width =1380
                    Height =225
                    TabIndex =8
                    Name ="Date"
                    ControlSource ="BeginDate"

                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10680
                    BorderColor =0
                    Name ="Line5"
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
                    Width =4680
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
            Height =1530
            Name ="ReportFooter"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Top =180
                    Width =465
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Sum"
                End
                Begin Label
                    FontItalic = NotDefault
                    Top =450
                    Width =420
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label30"
                    Caption ="Avg"
                End
                Begin Label
                    FontItalic = NotDefault
                    Top =720
                    Width =435
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label31"
                    Caption ="Min"
                End
                Begin Label
                    FontItalic = NotDefault
                    Top =990
                    Width =465
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label32"
                    Caption ="Max"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1860
                    Top =180
                    Width =840
                    Height =300
                    Name ="Sum Of SumOfPrice"
                    ControlSource ="=Sum([TotalSales])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_SumOfPrice"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1860
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =1
                    Name ="Avg Of SumOfPrice"
                    ControlSource ="=Avg([TotalSales])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_SumOfPrice"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1860
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =2
                    Name ="Min Of SumOfPrice"
                    ControlSource ="=Min([TotalSales])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_SumOfPrice"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1860
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =3
                    Name ="Max Of SumOfPrice"
                    ControlSource ="=Max([TotalSales])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_SumOfPrice"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2880
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =4
                    Name ="Sum Of 10 AM"
                    ControlSource ="=Sum([Sun])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2880
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =5
                    Name ="Avg Of 10 AM"
                    ControlSource ="=Avg([Sun])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2880
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =6
                    Name ="Min Of 10 AM"
                    ControlSource ="=Min([Sun])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2880
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =7
                    Name ="Max Of 10 AM"
                    ControlSource ="=Max([Sun])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3900
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =8
                    Name ="Sum Of 11 AM"
                    ControlSource ="=Sum([Mon])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3900
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =9
                    Name ="Avg Of 11 AM"
                    ControlSource ="=Avg([Mon])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3900
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =10
                    Name ="Min Of 11 AM"
                    ControlSource ="=Min([Mon])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3900
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =11
                    Name ="Max Of 11 AM"
                    ControlSource ="=Max([Mon])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =12
                    Name ="Sum Of 12 AM"
                    ControlSource ="=Sum([Tue])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =13
                    Name ="Avg Of 12 AM"
                    ControlSource ="=Avg([Tue])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =14
                    Name ="Min Of 12 AM"
                    ControlSource ="=Min([Tue])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4920
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =15
                    Name ="Max Of 12 AM"
                    ControlSource ="=Max([Tue])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =16
                    Name ="Sum Of 1 PM"
                    ControlSource ="=Sum([Wed])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =17
                    Name ="Avg Of 1 PM"
                    ControlSource ="=Avg([Wed])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =18
                    Name ="Min Of 1 PM"
                    ControlSource ="=Min([Wed])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =19
                    Name ="Max Of 1 PM"
                    ControlSource ="=Max([Wed])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =20
                    Name ="Sum Of 3 PM"
                    ControlSource ="=Sum([Thu])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =21
                    Name ="Avg Of 3 PM"
                    ControlSource ="=Avg([Thu])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =22
                    Name ="Min Of 3 PM"
                    ControlSource ="=Min([Thu])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =23
                    Name ="Max Of 3 PM"
                    ControlSource ="=Max([Thu])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7980
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =24
                    Name ="Sum Of 4 PM"
                    ControlSource ="=Sum([Fri])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7980
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =25
                    Name ="Avg Of 4 PM"
                    ControlSource ="=Avg([Fri])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7980
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =26
                    Name ="Min Of 4 PM"
                    ControlSource ="=Min([Fri])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7980
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =27
                    Name ="Max Of 4 PM"
                    ControlSource ="=Max([Fri])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =28
                    Name ="Sum Of 5 PM"
                    ControlSource ="=Sum([Sat])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Top =450
                    Width =720
                    Height =300
                    TabIndex =29
                    Name ="Avg Of 5 PM"
                    ControlSource ="=Avg([Sat])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Top =720
                    Width =720
                    Height =300
                    TabIndex =30
                    Name ="Min Of 5 PM"
                    ControlSource ="=Min([Sat])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9360
                    Top =990
                    Width =720
                    Height =300
                    TabIndex =31
                    Name ="Max Of 5 PM"
                    ControlSource ="=Max([Sat])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_5_PM"

                End
                Begin Line
                    Top =120
                    Width =10680
                    BorderColor =0
                    Name ="Line19"
                End
                Begin Line
                    Top =60
                    Width =10680
                    BorderColor =0
                    Name ="Line133"
                End
            End
        End
    End
End
CodeBehindForm
' See "17 Sales By Day - Related.cls"
