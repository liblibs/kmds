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
    Width =13860
    DatasheetFontHeight =10
    ItemSuffix =148
    Left =525
    Top =1230
    RecSrcDt = Begin
        0x27fd0c11acb6e240
    End
    RecordSource ="qrySalesByHourWithCustomerCount"
    Caption ="SalesByHour"
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
            ControlSource ="Date"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-30
                    Width =13890
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="Sales by Hour"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =480
                    Width =13786
                    Height =300
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Top =780
                    Width =13846
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
                    Left =60
                    Top =60
                    Width =1380
                    Height =300
                    Name ="Date_Label"
                    Caption ="Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =1470
                    Top =60
                    Width =1230
                    Height =300
                    Name ="SumOfPrice_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2940
                    Top =60
                    Width =660
                    Height =300
                    Name ="10 AM_Label"
                    Caption ="10 AM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl10_AM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =3660
                    Top =60
                    Width =660
                    Height =300
                    Name ="11 AM_Label"
                    Caption ="11 AM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl11_AM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =4440
                    Top =60
                    Width =660
                    Height =300
                    Name ="12 AM_Label"
                    Caption ="12 AM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl12_AM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =5220
                    Top =60
                    Width =660
                    Height =300
                    Name ="1 PM_Label"
                    Caption ="1 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl1_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =6000
                    Top =60
                    Width =660
                    Height =300
                    Name ="2 PM_Label"
                    Caption ="2 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl2_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =6780
                    Top =60
                    Width =660
                    Height =300
                    Name ="3 PM_Label"
                    Caption ="3 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl3_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =7500
                    Top =60
                    Width =660
                    Height =300
                    Name ="4 PM_Label"
                    Caption ="4 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl4_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =8280
                    Top =60
                    Width =660
                    Height =300
                    Name ="5 PM_Label"
                    Caption ="5 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl5_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9060
                    Top =60
                    Width =660
                    Height =300
                    Name ="6 PM_Label"
                    Caption ="6 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl6_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =9840
                    Top =60
                    Width =660
                    Height =300
                    Name ="7 PM_Label"
                    Caption ="7 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl7_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =10620
                    Top =60
                    Width =660
                    Height =300
                    Name ="8 PM_Label"
                    Caption ="8 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl8_PM_Label"
                End
                Begin Label
                    TextAlign =3
                    Left =11400
                    Top =60
                    Width =660
                    Height =300
                    Name ="9 PM_Label"
                    Caption ="9 PM"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl9_PM_Label"
                End
                Begin Line
                    Top =420
                    Width =13860
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Label
                    TextAlign =3
                    Left =12180
                    Top =60
                    Width =900
                    Height =300
                    Name ="Label146"
                    Caption ="Licensee"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =13170
                    Top =60
                    Width =690
                    Height =300
                    Name ="Label147"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
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
                    ControlSource ="SumOfPrice"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2760
                    Width =840
                    TabIndex =1
                    Name ="10 AM"
                    ControlSource ="10 AM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3540
                    Width =840
                    Height =225
                    TabIndex =2
                    Name ="11 AM"
                    ControlSource ="11 AM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4320
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="12 AM"
                    ControlSource ="12 AM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Width =840
                    Height =225
                    TabIndex =4
                    Name ="1 PM"
                    ControlSource ="1 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Width =840
                    Height =225
                    TabIndex =5
                    Name ="2 PM"
                    ControlSource ="2 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl2_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6660
                    Width =840
                    Height =225
                    TabIndex =6
                    Name ="3 PM"
                    ControlSource ="3 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Width =840
                    Height =225
                    TabIndex =7
                    Name ="4 PM"
                    ControlSource ="4 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Width =840
                    Height =225
                    TabIndex =8
                    Name ="5 PM"
                    ControlSource ="5 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Width =840
                    Height =225
                    TabIndex =9
                    Name ="6 PM"
                    ControlSource ="6 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl6_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Width =840
                    Height =225
                    TabIndex =10
                    Name ="7 PM"
                    ControlSource ="7 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl7_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10500
                    Width =840
                    Height =225
                    TabIndex =11
                    Name ="8 PM"
                    ControlSource ="8 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl8_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11280
                    Width =840
                    Height =225
                    TabIndex =12
                    Name ="9 PM"
                    ControlSource ="9 PM"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Ctl9_PM"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Width =1380
                    Height =225
                    TabIndex =13
                    Name ="Date"
                    ControlSource ="Date"

                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =13860
                    BorderColor =0
                    Name ="Line5"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12360
                    Width =720
                    Height =255
                    TabIndex =14
                    Name ="Licensee"
                    ControlSource ="Licensee"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13140
                    Width =720
                    Height =255
                    TabIndex =15
                    Name ="Count"
                    ControlSource ="Count"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter1"
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
                    Left =6720
                    Top =240
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
                    Width =4920
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
                    Left =6720
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
                    Left =60
                    Top =300
                    Width =11760
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
                    Left =240
                    Width =11580
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
            Height =1350
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
                    ControlSource ="=Sum([SumOfPrice])"
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
                    ControlSource ="=Avg([SumOfPrice])"
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
                    ControlSource ="=Min([SumOfPrice])"
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
                    ControlSource ="=Max([SumOfPrice])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_SumOfPrice"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2760
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =4
                    Name ="Sum Of 10 AM"
                    ControlSource ="=Sum([10 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2760
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =5
                    Name ="Avg Of 10 AM"
                    ControlSource ="=Avg([10 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2760
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =6
                    Name ="Min Of 10 AM"
                    ControlSource ="=Min([10 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2760
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =7
                    Name ="Max Of 10 AM"
                    ControlSource ="=Max([10 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_10_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3480
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =8
                    Name ="Sum Of 11 AM"
                    ControlSource ="=Sum([11 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3480
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =9
                    Name ="Avg Of 11 AM"
                    ControlSource ="=Avg([11 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3480
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =10
                    Name ="Min Of 11 AM"
                    ControlSource ="=Min([11 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3480
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =11
                    Name ="Max Of 11 AM"
                    ControlSource ="=Max([11 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_11_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =12
                    Name ="Sum Of 12 AM"
                    ControlSource ="=Sum([12 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =13
                    Name ="Avg Of 12 AM"
                    ControlSource ="=Avg([12 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =14
                    Name ="Min Of 12 AM"
                    ControlSource ="=Min([12 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =15
                    Name ="Max Of 12 AM"
                    ControlSource ="=Max([12 AM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_12_AM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =16
                    Name ="Sum Of 1 PM"
                    ControlSource ="=Sum([1 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =17
                    Name ="Avg Of 1 PM"
                    ControlSource ="=Avg([1 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =18
                    Name ="Min Of 1 PM"
                    ControlSource ="=Min([1 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =19
                    Name ="Max Of 1 PM"
                    ControlSource ="=Max([1 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_1_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =20
                    Name ="Sum Of 2 PM"
                    ControlSource ="=Sum([2 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_2_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =21
                    Name ="Avg Of 2 PM"
                    ControlSource ="=Avg([2 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_2_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =22
                    Name ="Min Of 2 PM"
                    ControlSource ="=Min([2 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_2_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =23
                    Name ="Max Of 2 PM"
                    ControlSource ="=Max([2 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_2_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6660
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =24
                    Name ="Sum Of 3 PM"
                    ControlSource ="=Sum([3 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6660
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =25
                    Name ="Avg Of 3 PM"
                    ControlSource ="=Avg([3 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6660
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =26
                    Name ="Min Of 3 PM"
                    ControlSource ="=Min([3 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6660
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =27
                    Name ="Max Of 3 PM"
                    ControlSource ="=Max([3 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_3_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =28
                    Name ="Sum Of 4 PM"
                    ControlSource ="=Sum([4 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =29
                    Name ="Avg Of 4 PM"
                    ControlSource ="=Avg([4 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =30
                    Name ="Min Of 4 PM"
                    ControlSource ="=Min([4 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7380
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =31
                    Name ="Max Of 4 PM"
                    ControlSource ="=Max([4 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_4_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =32
                    Name ="Sum Of 5 PM"
                    ControlSource ="=Sum([5 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =33
                    Name ="Avg Of 5 PM"
                    ControlSource ="=Avg([5 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =34
                    Name ="Min Of 5 PM"
                    ControlSource ="=Min([5 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =35
                    Name ="Max Of 5 PM"
                    ControlSource ="=Max([5 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_5_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =36
                    Name ="Sum Of 6 PM"
                    ControlSource ="=Sum([6 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_6_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =37
                    Name ="Avg Of 6 PM"
                    ControlSource ="=Avg([6 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_6_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =38
                    Name ="Min Of 6 PM"
                    ControlSource ="=Min([6 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_6_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8940
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =39
                    Name ="Max Of 6 PM"
                    ControlSource ="=Max([6 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_6_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =40
                    Name ="Sum Of 7 PM"
                    ControlSource ="=Sum([7 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_7_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =41
                    Name ="Avg Of 7 PM"
                    ControlSource ="=Avg([7 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_7_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =42
                    Name ="Min Of 7 PM"
                    ControlSource ="=Min([7 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_7_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =43
                    Name ="Max Of 7 PM"
                    ControlSource ="=Max([7 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_7_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10500
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =44
                    Name ="Sum Of 8 PM"
                    ControlSource ="=Sum([8 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_8_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10500
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =45
                    Name ="Avg Of 8 PM"
                    ControlSource ="=Avg([8 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_8_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10500
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =46
                    Name ="Min Of 8 PM"
                    ControlSource ="=Min([8 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_8_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10500
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =47
                    Name ="Max Of 8 PM"
                    ControlSource ="=Max([8 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_8_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11280
                    Top =180
                    Width =840
                    Height =300
                    TabIndex =48
                    Name ="Sum Of 9 PM"
                    ControlSource ="=Sum([9 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_9_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11280
                    Top =450
                    Width =840
                    Height =300
                    TabIndex =49
                    Name ="Avg Of 9 PM"
                    ControlSource ="=Avg([9 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Avg_Of_9_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11280
                    Top =720
                    Width =840
                    Height =300
                    TabIndex =50
                    Name ="Min Of 9 PM"
                    ControlSource ="=Min([9 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Min_Of_9_PM"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11280
                    Top =990
                    Width =840
                    Height =300
                    TabIndex =51
                    Name ="Max Of 9 PM"
                    ControlSource ="=Max([9 PM])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Max_Of_9_PM"

                End
                Begin Line
                    Top =120
                    Width =13860
                    BorderColor =0
                    Name ="Line19"
                End
                Begin Line
                    Top =60
                    Width =13860
                    BorderColor =0
                    Name ="Line133"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12300
                    Top =180
                    Width =840
                    Height =330
                    TabIndex =52
                    Name ="Text138"
                    ControlSource ="=Sum([Licensee])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12300
                    Top =450
                    Width =840
                    Height =330
                    TabIndex =53
                    Name ="Text139"
                    ControlSource ="=Avg([Licensee])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12300
                    Top =720
                    Width =840
                    Height =330
                    TabIndex =54
                    Name ="Text140"
                    ControlSource ="=Min([Licensee])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12300
                    Top =990
                    Width =840
                    Height =330
                    TabIndex =55
                    Name ="Text141"
                    ControlSource ="=Max([Licensee])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13260
                    Top =180
                    Width =600
                    Height =360
                    TabIndex =56
                    Name ="Text142"
                    ControlSource ="=Sum([Count])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13260
                    Top =450
                    Width =600
                    Height =360
                    TabIndex =57
                    Name ="Text143"
                    ControlSource ="=Avg([Count])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13260
                    Top =720
                    Width =600
                    Height =360
                    TabIndex =58
                    Name ="Text144"
                    ControlSource ="=Min([Count])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13260
                    Top =990
                    Width =600
                    Height =360
                    TabIndex =59
                    Name ="Text145"
                    ControlSource ="=Max([Count])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
            End
        End
    End
End
CodeBehindForm
' See "16 Sales By Hour.cls"
