Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14340
    DatasheetFontHeight =10
    ItemSuffix =69
    Left =1710
    Top =2670
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x7a4de71b70d0e240
    End
    RecordSource ="qryPurchaseOrderCreate"
    Caption ="PurchaseOrderCreate"
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
            Height =1920
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =660
                    Width =900
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =1
                    Left =1020
                    Top =660
                    Width =9594
                    Height =510
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =4260
                    Height =510
                    FontSize =20
                    Name ="Label51"
                    Caption ="Purchase Order Create"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =14325
                    BorderColor =12632256
                    Name ="Line54"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =14325
                    BorderColor =12632256
                    Name ="Line55"
                End
                Begin Line
                    BorderWidth =3
                    Top =1680
                    Width =14325
                    BorderColor =12632256
                    Name ="Line57"
                End
                Begin Label
                    Visible = NotDefault
                    Left =240
                    Top =1260
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
            End
        End
        Begin PageHeader
            Height =405
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =2940
                    Top =60
                    Width =690
                    Height =300
                    Name ="SizeID_Label"
                    Caption ="SizeID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =4380
                    Top =60
                    Width =810
                    Height =300
                    Name ="Onhand_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5220
                    Top =60
                    Width =825
                    Height =300
                    Name ="Ordered_Label"
                    Caption ="Ordered"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6000
                    Top =60
                    Width =600
                    Height =300
                    Name ="Reserved_Label"
                    Caption ="Rsrvd"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6780
                    Top =60
                    Width =555
                    Height =300
                    Name ="Avail_Label"
                    Caption ="Avail"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7680
                    Top =60
                    Width =450
                    Height =300
                    Name ="OrderAt_Label"
                    Caption ="Min"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8340
                    Top =60
                    Width =495
                    Height =300
                    Name ="OrderTo_Label"
                    Caption ="Max"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8940
                    Top =60
                    Width =660
                    Height =300
                    Name ="SOQRaw_Label"
                    Caption ="RSOQ"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =9630
                    Top =60
                    Width =765
                    Height =300
                    Name ="OrderMultiple_Label"
                    Caption ="OrderQ"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3780
                    Top =60
                    Width =510
                    Height =300
                    Name ="SOQ_Label"
                    Caption ="SOQ"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =10440
                    Top =60
                    Width =675
                    Height =300
                    Name ="UnitsPerCase_Label"
                    Caption ="UCase"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =11220
                    Top =60
                    Width =555
                    Height =300
                    Name ="SPrice_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =12000
                    Top =60
                    Width =870
                    Height =300
                    Name ="ExtPrice_Label"
                    Caption ="ExtPrice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =12960
                    Top =60
                    Width =480
                    Height =300
                    Name ="Cost_Label"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =13440
                    Top =60
                    Width =795
                    Height =300
                    Name ="ExtCost_Label"
                    Caption ="ExtCost"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3000
                    Top =60
                    Width =11250
                    Name ="Line58"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3000
                    Top =360
                    Width =11250
                    Name ="Line60"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3000
                    Top =390
                    Width =11250
                    Name ="Line61"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3000
                    Width =11250
                    Name ="Line66"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            Name ="Detail"
            Begin
                Begin TextBox
                    Top =60
                    Width =600
                    Height =225
                    ColumnWidth =765
                    Name ="ItemID"
                    ControlSource ="ItemID"

                End
                Begin TextBox
                    Left =660
                    Top =60
                    Width =2340
                    Height =225
                    ColumnWidth =3225
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnCount =2
                    ListRows =20
                    Left =3060
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =765
                    TabIndex =2
                    Name ="SizeID"
                    ControlSource ="SizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.ID, Size.Description FROM [Size] ORDER BY Size.ID; "
                    ColumnWidths ="0;1080"

                End
                Begin TextBox
                    TextAlign =3
                    Left =4500
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =900
                    TabIndex =3
                    Name ="Onhand"
                    ControlSource ="Onhand"

                End
                Begin TextBox
                    TextAlign =3
                    Left =5245
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =930
                    TabIndex =4
                    Name ="Ordered"
                    ControlSource ="Ordered"

                End
                Begin TextBox
                    TextAlign =3
                    Left =5990
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =1035
                    TabIndex =5
                    Name ="Reserved"
                    ControlSource ="Reserved"

                End
                Begin TextBox
                    TextAlign =3
                    Left =6735
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =630
                    TabIndex =6
                    Name ="Avail"
                    ControlSource ="Avail"

                End
                Begin TextBox
                    TextAlign =3
                    Left =7480
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =885
                    TabIndex =7
                    Name ="OrderAt"
                    ControlSource ="OrderAt"

                End
                Begin TextBox
                    TextAlign =3
                    Left =8225
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =930
                    TabIndex =8
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                End
                Begin TextBox
                    TextAlign =3
                    Left =8970
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =1005
                    TabIndex =9
                    Name ="SOQRaw"
                    ControlSource ="SOQRaw"

                End
                Begin TextBox
                    TextAlign =3
                    Left =9715
                    Top =60
                    Width =660
                    Height =225
                    ColumnWidth =1455
                    TabIndex =10
                    Name ="OrderMultiple"
                    ControlSource ="OrderMultiple"

                End
                Begin TextBox
                    TextAlign =3
                    Left =3780
                    Top =60
                    Width =540
                    Height =225
                    ColumnWidth =585
                    FontWeight =700
                    TabIndex =11
                    Name ="SOQ"
                    ControlSource ="SOQ"

                End
                Begin TextBox
                    TextAlign =3
                    Left =10485
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =12
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"

                End
                Begin TextBox
                    TextAlign =3
                    Left =11230
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =13
                    Name ="SPrice"
                    ControlSource ="SPrice"

                End
                Begin TextBox
                    TextAlign =3
                    Left =11940
                    Top =60
                    Width =270
                    Height =225
                    ColumnWidth =1095
                    TabIndex =14
                    Name ="SPriceStar"
                    ControlSource ="SPriceStar"

                End
                Begin TextBox
                    TextAlign =3
                    Left =12265
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =15
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"

                End
                Begin TextBox
                    TextAlign =3
                    Left =12925
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =16
                    Name ="Cost"
                    ControlSource ="Cost"

                End
                Begin TextBox
                    TextAlign =3
                    Left =13645
                    Top =60
                    Width =660
                    Height =225
                    TabIndex =17
                    Name ="ExtCost"
                    ControlSource ="ExtCost"

                End
                Begin Line
                    Top =420
                    Width =14340
                    Name ="Line65"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2700
                    Top =60
                    Width =300
                    TabIndex =18
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"

                End
            End
        End
        Begin PageFooter
            Height =660
            Name ="PageFooterSection"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =14325
                    BorderColor =12632256
                    Name ="Line63"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =14325
                    BorderColor =12632256
                    Name ="Line64"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Top =360
                    Width =8640
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
                    Left =5640
                    Top =360
                    Width =8640
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
                    Top =60
                    Width =8640
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
                    Left =5640
                    Top =60
                    Width =8640
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
                    Top =360
                    Width =14280
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
                    Top =60
                    Width =14280
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
            Height =525
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Width =13860
                    Height =225
                    Name ="Text2"
                    ControlSource ="=\"Summary for \" & \"'VendorID' = \" & \" \" & [VendorID] & \" (\" & Count(*) &"
                        " \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =2700
                    Top =240
                    Width =525
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3420
                    Top =270
                    Width =900
                    Height =225
                    TabIndex =1
                    Name ="Sum Of SOQ"
                    ControlSource ="=Sum([SOQ])"
                    EventProcPrefix ="Sum_Of_SOQ"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =10920
                    Top =270
                    Width =900
                    Height =225
                    TabIndex =2
                    Name ="Sum Of ExtPrice"
                    ControlSource ="=Sum([ExtPrice])"
                    EventProcPrefix ="Sum_Of_ExtPrice"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =13260
                    Top =300
                    Width =1035
                    Height =225
                    TabIndex =3
                    Name ="Sum Of ExtCost"
                    ControlSource ="=Sum([ExtCost])"
                    EventProcPrefix ="Sum_Of_ExtCost"

                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrderCreate.cls"
