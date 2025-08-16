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
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =132
    Top =1080
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x1ef574cc6e28e440
    End
    RecordSource ="qryRestockList"
    Caption ="End of Day Item Restock List"
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
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10665
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="End of Day Item Restock List"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =10725
                    LayoutCachedHeight =600
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10680
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1320
                    Width =10680
                    BorderColor =12632256
                    Name ="Line53"
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =960
                    Width =10620
                    Height =300
                    ColumnOrder =1
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text125"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =60
                    LayoutCachedTop =960
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Top =660
                    Width =10620
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text126"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =960
                End
            End
        End
        Begin PageHeader
            Height =1080
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =0
                    Left =360
                    Top =420
                    Width =555
                    Height =330
                    Name ="Department_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =360
                    LayoutCachedTop =420
                    LayoutCachedWidth =915
                    LayoutCachedHeight =750
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10620
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =60
                    Top =60
                    Width =10620
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =840
                    Width =10620
                    Name ="Line56"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =840
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =840
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =60
                    Top =900
                    Width =10620
                    Name ="Line57"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =900
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =900
                End
                Begin Label
                    TextAlign =2
                    Left =4980
                    Top =120
                    Width =1140
                    Height =660
                    Name ="Label114"
                    Caption ="Restock Amount"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4980
                    LayoutCachedTop =120
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =780
                End
                Begin Label
                    Left =2040
                    Top =480
                    Width =900
                    Height =300
                    Name ="Label116"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =480
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =780
                End
                Begin Label
                    TextAlign =3
                    Left =1380
                    Top =480
                    Width =525
                    Height =330
                    Name ="Label120"
                    Caption ="Size"
                    LayoutCachedLeft =1380
                    LayoutCachedTop =480
                    LayoutCachedWidth =1905
                    LayoutCachedHeight =810
                End
                Begin Label
                    TextAlign =0
                    Left =7080
                    Top =480
                    Width =930
                    Height =330
                    Name ="Label121"
                    Caption ="Category"
                    LayoutCachedLeft =7080
                    LayoutCachedTop =480
                    LayoutCachedWidth =8010
                    LayoutCachedHeight =810
                End
                Begin Label
                    TextAlign =0
                    Left =6000
                    Top =480
                    Width =915
                    Height =300
                    Name ="Label123"
                    Caption ="OnShelf"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =480
                    LayoutCachedWidth =6915
                    LayoutCachedHeight =780
                End
                Begin Label
                    TextAlign =0
                    Left =8460
                    Top =450
                    Width =2295
                    Height =330
                    Name ="Label128"
                    Caption ="Discrepency/Comments"
                    LayoutCachedLeft =8460
                    LayoutCachedTop =450
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =780
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =120
                    Top =60
                    Width =1320
                    Height =285
                    FontSize =10
                    Name ="Text82"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Left =2040
                    Top =60
                    Width =3195
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Text83"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =5235
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5160
                    Top =60
                    Width =795
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="RestockAmount"
                    ControlSource ="RestockAmount"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =5955
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7140
                    Top =60
                    Width =1395
                    Height =225
                    TabIndex =3
                    Name ="SubCategory.Description"
                    ControlSource ="Category"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="SubCategory_Description"

                    LayoutCachedLeft =7140
                    LayoutCachedTop =60
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =900
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="Size"
                    ControlSource ="Size"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6060
                    Top =60
                    Width =720
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6060
                    LayoutCachedTop =60
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =330
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =420
                    Width =10620
                    Name ="Line131"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =420
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =420
                End
            End
        End
        Begin PageFooter
            Height =1260
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

                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =690
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

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =345
                End
                Begin Label
                    OldBorderStyle =1
                    Left =180
                    Top =780
                    Width =4725
                    Height =420
                    Name ="Label129"
                    Caption ="Completed by: ___________________________"
                    LayoutCachedLeft =180
                    LayoutCachedTop =780
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OldBorderStyle =1
                    Left =5460
                    Top =780
                    Width =5055
                    Height =420
                    Name ="Label130"
                    Caption ="Time Completed : ___________________________"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =780
                    LayoutCachedWidth =10515
                    LayoutCachedHeight =1200
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =780
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextFontFamily =34
                    Top =120
                    Width =1620
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Total"
                    FontName ="Arial Narrow"
                    LayoutCachedTop =120
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =390
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =960
                    Top =120
                    Width =1980
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =960
                    LayoutCachedTop =120
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =3
                    Top =480
                    Width =10680
                    BorderColor =12632256
                    Name ="Line78"
                    LayoutCachedTop =480
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =5520
                    Top =120
                    Width =720
                    FontWeight =700
                    Name ="Text110"
                    ControlSource ="=Sum([RestockAmount])"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7560
                    Top =180
                    Width =900
                    FontWeight =700
                    TabIndex =1
                    Name ="Text111"
                    ControlSource ="=Sum([Dollars])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =180
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD126ItemRestockListNOSubtotals.cls"
