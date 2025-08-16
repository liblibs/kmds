Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =53
    Left =585
    Top =1320
    OnNoData ="[Event Procedure]"
    OrderBy ="ItemID"
    RecSrcDt = Begin
        0x12b0579599e8e240
    End
    RecordSource ="qryAuditReport"
    Caption ="qryAudit"
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
            Height =570
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =5205
                    Height =510
                    FontSize =20
                    Name ="Label22"
                    Caption ="Onhand Change Audit Report"
                End
            End
        End
        Begin PageHeader
            Height =375
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =840
                    Height =300
                    Name ="ItemID_Label"
                    Caption ="ItemID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =60
                    Width =825
                    Height =300
                    Name ="OnShelf_Label"
                    Caption ="OnShelf"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4680
                    Top =60
                    Width =630
                    Height =300
                    Name ="PhysicalCount_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9300
                    Top =60
                    Width =540
                    Height =300
                    Name ="Price_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7920
                    Top =60
                    Width =540
                    Height =300
                    Name ="Cost_Label"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9840
                    Top =60
                    Width =840
                    Height =300
                    Name ="ExtPrice_Label"
                    Caption ="ExtPrice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8460
                    Top =60
                    Width =840
                    Height =300
                    Name ="ExtCost_Label"
                    Caption ="ExtCost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6960
                    Top =60
                    Width =900
                    Height =300
                    Name ="VarianceAmount_Label"
                    Caption ="Variance"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =360
                    Width =10620
                    Name ="Line25"
                End
                Begin Label
                    TextAlign =3
                    Left =5640
                    Top =60
                    Width =1290
                    Height =300
                    Name ="Label48"
                    Caption ="F - B - D - O"
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
                    IMESentenceMode =3
                    Left =60
                    Width =1260
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4005
                    Width =600
                    TabIndex =1
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4005
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4710
                    Width =600
                    TabIndex =2
                    Name ="PhysicalCount"
                    ControlSource ="PhysicalCount"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4710
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7260
                    Width =600
                    Height =225
                    TabIndex =3
                    Name ="Var"
                    ControlSource ="Var"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7260
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9300
                    Width =540
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9300
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7920
                    Width =540
                    TabIndex =5
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9840
                    Width =840
                    TabIndex =6
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9840
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8460
                    Width =840
                    TabIndex =7
                    Name ="ExtCost"
                    ControlSource ="ExtCost"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8460
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1320
                    Width =2220
                    TabIndex =8
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1320
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3540
                    Width =420
                    TabIndex =9
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3540
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =240
                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10680
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5580
                    Width =330
                    TabIndex =10
                    Name ="CountFront"
                    ControlSource ="CountFront"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5580
                    LayoutCachedWidth =5910
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5910
                    Width =330
                    TabIndex =11
                    Name ="CountBack"
                    ControlSource ="CountBack"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5910
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6270
                    Width =330
                    TabIndex =12
                    Name ="CountDisplay"
                    ControlSource ="CountDisplay"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6270
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6630
                    Width =330
                    TabIndex =13
                    Name ="CountOverflow"
                    ControlSource ="CountOverflow"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6630
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =240
                End
                Begin Line
                    Left =5460
                    Width =0
                    Height =240
                    Name ="Line50"
                End
                Begin Line
                    Left =7020
                    Width =0
                    Height =240
                    Name ="Line51"
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
                    Width =4920
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
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
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
                    Width =4920
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
                    Width =10560
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
                    Width =10560
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
            Height =1020
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3540
                    Top =180
                    Width =840
                    Height =285
                    Name ="Text30"
                    ControlSource ="=Sum([OnShelf])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4470
                    Top =180
                    Width =840
                    Height =285
                    TabIndex =1
                    Name ="Text31"
                    ControlSource ="=Sum([PhysicalCount])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6960
                    Top =180
                    Width =960
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text32"
                    ControlSource ="=Sum([Var])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =180
                    Width =960
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text35"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8040
                    Top =180
                    Width =840
                    Height =285
                    TabIndex =4
                    Name ="Text36"
                    ControlSource ="=Sum([ExtCost])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin Label
                    TextAlign =3
                    Left =60
                    Top =180
                    Width =2580
                    Height =300
                    Name ="Label39"
                    Caption ="Total"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Width =10620
                    Name ="Line43"
                End
                Begin Label
                    TextAlign =3
                    Left =60
                    Top =600
                    Width =4470
                    Height =300
                    Name ="Label52"
                    Caption ="F - B - D - 0 = Front - Back - Display - Overflow"
                End
            End
        End
    End
End
CodeBehindForm
' See "Audit.cls"
