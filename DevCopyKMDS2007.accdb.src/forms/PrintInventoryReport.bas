Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11820
    DatasheetFontHeight =10
    ItemSuffix =108
    Left =3660
    Top =1500
    Right =15810
    Bottom =9285
    HelpContextId =4300
    Filter ="[ReportName]='52 Purchases By Vendor By Item'"
    RecSrcDt = Begin
        0x9ce51521e9a8e240
    End
    RecordSource ="InventoryReportParameters"
    Caption =" KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Softwa"
        "re ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS S"
        "oftware ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  K"
        "MDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software"
        " ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Sof"
        "tware ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionGroup
            SpecialEffect =3
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =2
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CustomControl
            SpecialEffect =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =7800
            BackColor =-2147483633
            Name ="Detail"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =5490
                    Top =1140
                    Width =1770
                    TabIndex =4
                    Name ="cboSalesDepartmentIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [SalesDepartment].[ID], [SalesDepartment].[Description] FROM SalesDepartm"
                        "ent ORDER BY [ID]; "
                    ColumnWidths ="360;1950"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =1140
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4080
                            Top =1140
                            Width =1350
                            Height =240
                            Name ="From Sales Department Number_Label"
                            Caption ="Sales Department"
                            EventProcPrefix ="From_Sales_Department_Number_Label"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =1140
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =1380
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =7440
                    Top =1140
                    Width =1770
                    TabIndex =5
                    Name ="cboSalesDepartmentIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [SalesDepartment].[ID], [SalesDepartment].[Description] FROM SalesDepartm"
                        "ent ORDER BY [ID]; "
                    ColumnWidths ="360;1950"
                    DefaultValue ="1"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =1380
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8535
                    Top =4845
                    TabIndex =18
                    Name ="chkPrintZeroOnhands"
                    DefaultValue ="True"

                    LayoutCachedLeft =8535
                    LayoutCachedTop =4845
                    LayoutCachedWidth =8795
                    LayoutCachedHeight =5085
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6075
                            Top =4785
                            Width =2400
                            Height =240
                            Name ="Label8"
                            Caption ="Print Items with Onhands = 0"
                            LayoutCachedLeft =6075
                            LayoutCachedTop =4785
                            LayoutCachedWidth =8475
                            LayoutCachedHeight =5025
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8535
                    Top =4545
                    TabIndex =17
                    Name ="chkCost"
                    DefaultValue ="False"

                    LayoutCachedLeft =8535
                    LayoutCachedTop =4545
                    LayoutCachedWidth =8795
                    LayoutCachedHeight =4785
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =7050
                            Top =4485
                            Width =1425
                            Height =240
                            Name ="Label11"
                            Caption ="Print Cost"
                            LayoutCachedLeft =7050
                            LayoutCachedTop =4485
                            LayoutCachedWidth =8475
                            LayoutCachedHeight =4725
                        End
                    End
                End
                Begin OptionGroup
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =4020
                    Top =360
                    Width =5280
                    Height =6060
                    Name ="Frame18"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =360
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =6420
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =255
                            Left =4140
                            Top =240
                            Width =2595
                            Height =240
                            BackColor =-2147483633
                            Name ="Label19"
                            Caption ="Select range of inventory to print.."
                            LayoutCachedLeft =4140
                            LayoutCachedTop =240
                            LayoutCachedWidth =6735
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5220
                    Top =540
                    Width =3960
                    TabIndex =2
                    BackColor =39423
                    Name ="txtReportName"
                    ControlSource ="ReportName"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =540
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4080
                            Top =540
                            Width =1080
                            Height =240
                            Name ="Label21"
                            Caption ="Report Name:"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =540
                            LayoutCachedWidth =5160
                            LayoutCachedHeight =780
                        End
                    End
                End
                Begin Label
                    OverlapFlags =255
                    Left =4860
                    Top =840
                    Width =1980
                    Height =240
                    Name ="Label22"
                    Caption ="Select range From"
                    LayoutCachedLeft =4860
                    LayoutCachedTop =840
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =255
                    Left =7440
                    Top =840
                    Width =840
                    Height =240
                    Name ="Label23"
                    Caption ="To"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =840
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1080
                End
                Begin ListBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =18
                    Left =240
                    Top =660
                    Width =3600
                    Height =6720
                    TabIndex =1
                    Name ="lstReportName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT InventoryReportParameters.ReportName, InventoryReportParameters.QueryName"
                        ", InventoryReportParameters.RangeSalesDepartments, InventoryReportParameters.Ran"
                        "geVendor, InventoryReportParameters.RangeLocations, InventoryReportParameters.Ra"
                        "ngeItemID, InventoryReportParameters.RangeDates, InventoryReportParameters.Print"
                        "Onhands, InventoryReportParameters.PrintCost, InventoryReportParameters.PrintZer"
                        "oOnhands, InventoryReportParameters.ChartQueryName, InventoryReportParameters.Ra"
                        "ngePriceGroupID, InventoryReportParameters.NumericRange, InventoryReportParamete"
                        "rs.NumericRangeLeadThru, InventoryReportParameters.ChartQueryName, InventoryRepo"
                        "rtParameters.RangeSalesDepartments, InventoryReportParameters.RangeClubs, Invent"
                        "oryReportParameters.RangeSubCategories FROM InventoryReportParameters ORDER BY I"
                        "nventoryReportParameters.ReportName; "
                    ColumnWidths ="2430;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =660
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =7380
                    Begin
                        Begin Label
                            SpecialEffect =3
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =180
                            Top =300
                            Width =3720
                            Height =7140
                            Name ="Label27"
                            Caption ="Select the report to print by clicking on it"
                            LayoutCachedLeft =180
                            LayoutCachedTop =300
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =20
                    ListWidth =2310
                    Left =5490
                    Top =1740
                    Width =1770
                    TabIndex =6
                    Name ="cboLocationBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Location FROM Item GROUP BY Item.Location ORDER BY Item.Location; "
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =1740
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1980
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4080
                            Top =1740
                            Width =1350
                            Height =240
                            Name ="Label29"
                            Caption ="Locations"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =1740
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ListRows =20
                    ListWidth =2310
                    Left =7440
                    Top =1740
                    Width =1770
                    TabIndex =7
                    Name ="cboLocationEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Location FROM Item GROUP BY Item.Location ORDER BY Item.Location; "
                    ColumnWidths ="1440"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =1740
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =1980
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =5490
                    Top =2040
                    Width =1770
                    TabIndex =8
                    Name ="cboVendorIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name, Vendor.POMethodDefault FROM Vendor INNER JOIN Ite"
                        "m ON Vendor.ID = Item.VendorID GROUP BY Vendor.ID, Vendor.Name, Vendor.POMethodD"
                        "efault HAVING (((Vendor.POMethodDefault)<>-1)) ORDER BY Vendor.ID; "
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =2040
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =2280
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4080
                            Top =2040
                            Width =1350
                            Height =240
                            Name ="Label32"
                            Caption ="Vendors"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =2040
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =2280
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =7440
                    Top =2040
                    Width =1770
                    TabIndex =9
                    Name ="cboVendorIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name, Vendor.POMethodDefault FROM Vendor INNER JOIN Ite"
                        "m ON Vendor.ID = Item.VendorID GROUP BY Vendor.ID, Vendor.Name, Vendor.POMethodD"
                        "efault HAVING (((Vendor.POMethodDefault)<>-1)) ORDER BY Vendor.ID; "
                    ColumnWidths ="0;2880"
                    DefaultValue ="1"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =2040
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =2280
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =8535
                    Top =4245
                    TabIndex =16
                    Name ="chkPrintOnhands"
                    DefaultValue ="True"

                    LayoutCachedLeft =8535
                    LayoutCachedTop =4245
                    LayoutCachedWidth =8795
                    LayoutCachedHeight =4485
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =6075
                            Top =4185
                            Width =2400
                            Height =240
                            Name ="Label37"
                            Caption ="Print Onhands"
                            LayoutCachedLeft =6075
                            LayoutCachedTop =4185
                            LayoutCachedWidth =8475
                            LayoutCachedHeight =4425
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =5490
                    Top =2640
                    Width =1770
                    TabIndex =10
                    Name ="cboItemIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description FROM Item ORDER BY Item.ID; "
                    ColumnWidths ="1080;2880"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =2640
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =2880
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4080
                            Top =2640
                            Width =1350
                            Height =240
                            Name ="Label39"
                            Caption ="Items"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =2640
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =2880
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =7440
                    Top =2640
                    Width =1770
                    TabIndex =11
                    Name ="cboItemIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description FROM Item ORDER BY Item.ID; "
                    ColumnWidths ="1080;2880"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =2640
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =2880
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =86
                    Left =4020
                    Top =6600
                    Width =606
                    Height =546
                    TabIndex =19
                    ForeColor =0
                    Name ="cmdPrintPreview"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"
                    UnicodeAccessKey =118

                    LayoutCachedLeft =4020
                    LayoutCachedTop =6600
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =7146
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4020
                            Top =7200
                            Width =1020
                            Height =240
                            Name ="Label41"
                            Caption ="Print Pre&view"
                            LayoutCachedLeft =4020
                            LayoutCachedTop =7200
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =80
                    Left =5100
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =20
                    ForeColor =0
                    Name ="cmdPrint"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad00000000000dadd0888888888080da ,
                        0x000000000000080d0888888bbb88000a088888877788080d0000000000000880 ,
                        0x0888888888808080d000000000080800ad0ffffffff08080dad0f00000f0000a ,
                        0xada0ffffffff0daddada0f00000f0adaadad0ffffffff0addadad000000000da ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Print Report"
                    UnicodeAccessKey =80

                    LayoutCachedLeft =5100
                    LayoutCachedTop =6600
                    LayoutCachedWidth =5676
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =5160
                            Top =7200
                            Width =405
                            Height =240
                            Name ="Label9"
                            Caption ="&Print"
                            LayoutCachedLeft =5160
                            LayoutCachedTop =7200
                            LayoutCachedWidth =5565
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =67
                    Left =10260
                    Top =6540
                    Width =666
                    Height =606
                    TabIndex =21
                    ForeColor =0
                    Name ="cmdClose"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =10260
                    LayoutCachedTop =6540
                    LayoutCachedWidth =10926
                    LayoutCachedHeight =7146
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =10980
                            Top =6540
                            Width =585
                            Height =405
                            Name ="Label65"
                            Caption ="&Close Form"
                            LayoutCachedLeft =10980
                            LayoutCachedTop =6540
                            LayoutCachedWidth =11565
                            LayoutCachedHeight =6945
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =9360
                    Top =1140
                    Width =2220
                    Height =2040
                    TabIndex =3
                    BoundColumn =1
                    Name ="cmbOrderBy"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT InventoryReportOrderBys.ReportName, InventoryReportOrderBys.OrderByText, "
                        "InventoryReportOrderBys.ID FROM InventoryReportOrderBys WHERE (((InventoryReport"
                        "OrderBys.ReportName)=[Forms]![PrintInventoryReport]![txtReportName])) ORDER BY I"
                        "nventoryReportOrderBys.ID; "
                    ColumnWidths ="0;2160;0"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =1140
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =3180
                    Begin
                        Begin Label
                            SpecialEffect =3
                            OldBorderStyle =1
                            OverlapFlags =127
                            Left =9300
                            Top =360
                            Width =2220
                            Height =720
                            Name ="Order by_Label"
                            Caption ="Orderby  (You only need to change this if you want to print the report in a diff"
                                "erent order!)"
                            EventProcPrefix ="Order_by_Label"
                            LayoutCachedLeft =9300
                            LayoutCachedTop =360
                            LayoutCachedWidth =11520
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =180
                    Width =11700
                    Height =7380
                    Name ="Box45"
                    LayoutCachedLeft =120
                    LayoutCachedTop =180
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =7560
                End
                Begin CustomControl
                    Visible = NotDefault
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =5490
                    Top =3240
                    Width =1770
                    Height =300
                    AutoActivate =1
                    TabIndex =12
                    Name ="dtpDateRangeFrom"
                    OleData = Begin
                        0x00120000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000009e1bdd20c487d1118be30000 ,
                        0xf8754da1000000000000000000000000d09a20e43ef8c5010a00000080020000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010200000003000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000700000093000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000600000014000000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e0002010500000004000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000003000000b8000000 ,
                        0x0000000009000000fdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfefffffffefffffffeffffff07000000ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000009e1bdd20c487d1118be30000 ,
                        0xf8754da100000000000000000000000050997a3306e9cc010600000080020000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010200000003000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000700000093000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000600000014000000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e0002010500000004000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000003000000b8000000 ,
                        0x00000000ffffffffffffffff05000000fdfffffffefffffffeffffff07000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff02000000feffffff0400000005000000fefffffffeffffff ,
                        0x0800000009000000feffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff03004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000010000005c000000 ,
                        0x0000000002004f006c0065005000720065007300300030003000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000018000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000024000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000400000001000000ffffffff000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000036353141 ,
                        0x383934302d383743352d313164312d384245332d303030304638373534444131 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff9e1bdd20c487d1118be30000f8754da12143341208000000320c0000 ,
                        0x1102000041287f0e0000060020000000d4070900060004000000000000000000 ,
                        0x0f270c0005001f00000000000000000041060100010001000000000000000000 ,
                        0x080000000000000001efcdab0000050054dc230006000000ffffffffffffffff ,
                        0xc4b65d76000000001fdeecbd0100050061dc23000352e30b918fce119de300aa ,
                        0x004bb851010000009001444201000d4d532053616e73205365726966ffffffff ,
                        0xffffffff0100000200000000000000000000000000000000ffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff0100feff030a0000ffffffff9e1bdd20c487d1118be30000f8754da1 ,
                        0x310000004d6963726f736f6674204461746520616e642054696d65205069636b ,
                        0x657220436f6e74726f6c20362e3020285350342900150000004d53436f6d4374 ,
                        0x6c322e44545069636b65722e3200150000004d53436f6d43746c322e44545069 ,
                        0x636b65722e3200f439b271000000000000000000000000ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff00000000
                    End
                    OnUpdated ="[Event Procedure]"
                    OLEClass ="DTPicker"
                    ControlSource ="DateFrom"
                    Class ="MSComCtl2.DTPicker.2"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =3240
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =3540
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4080
                            Top =3240
                            Width =1350
                            Height =285
                            Name ="lblDateRangeFrom"
                            Caption ="Date Range"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =3240
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =3525
                        End
                    End
                End
                Begin CustomControl
                    Visible = NotDefault
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =7320
                    Top =3240
                    Width =1770
                    Height =300
                    AutoActivate =1
                    TabIndex =13
                    Name ="dtpDateRangeTo"
                    OleData = Begin
                        0x00120000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000009e1bdd20c487d1118be30000 ,
                        0xf8754da1000000000000000000000000d09a20e43ef8c5010a00000080020000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010200000003000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000600000093000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000900000014000000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e0002010500000004000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000003000000b8000000 ,
                        0x0000000009000000fdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfefffffffefffffffeffffff07000000ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000009e1bdd20c487d1118be30000 ,
                        0xf8754da100000000000000000000000060c07a3306e9cc010600000080020000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010200000003000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000600000093000000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000900000014000000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e0002010500000004000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000003000000b8000000 ,
                        0x00000000ffffffffffffffff05000000fdfffffffefffffffeffffff07000000 ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff02000000feffffff0400000005000000feffffff07000000 ,
                        0x08000000fefffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff03004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000010000005c000000 ,
                        0x0000000002004f006c0065005000720065007300300030003000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000018000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000024000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000400000001000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000005c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000036353141 ,
                        0x383934302d383743352d313164312d384245332d303030304638373534444131 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000009e1bdd20c487d1118be30000f8754da12143341208000000320c0000 ,
                        0x1102000041287f0e0000060020000000d4070900060004000000000000000000 ,
                        0x0f270c0005001f00000000000000000041060100010001000000000000000000 ,
                        0x080000000000000001efcdab0000050054dc230006000000ffffffffffffffff ,
                        0xc4b65d76000000001fdeecbd0100050061dc23000352e30b918fce119de300aa ,
                        0x004bb851010000009001444201000d4d532053616e7320536572696600000000 ,
                        0x000000000100feff030a0000ffffffff9e1bdd20c487d1118be30000f8754da1 ,
                        0x310000004d6963726f736f6674204461746520616e642054696d65205069636b ,
                        0x657220436f6e74726f6c20362e3020285350342900150000004d53436f6d4374 ,
                        0x6c322e44545069636b65722e3200150000004d53436f6d43746c322e44545069 ,
                        0x636b65722e3200f439b271000000000000000000000000ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff0100000200000000000000000000000000000000ffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff00000000
                    End
                    OnUpdated ="[Event Procedure]"
                    OLEClass ="DTPicker"
                    ControlSource ="DateTo"
                    Class ="MSComCtl2.DTPicker.2"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =3240
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =3540
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5520
                    Top =3540
                    Width =1770
                    TabIndex =14
                    Name ="txtDateRangeFrom"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =3540
                    LayoutCachedWidth =7290
                    LayoutCachedHeight =3780
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4800
                            Top =3540
                            Width =645
                            Height =240
                            Name ="Label49"
                            Caption ="Dates"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =3540
                            LayoutCachedWidth =5445
                            LayoutCachedHeight =3780
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7320
                    Top =3540
                    Width =1770
                    TabIndex =15
                    Name ="txtDateRangeTo"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =3540
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =3780
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =72
                    Left =8580
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =22
                    ForeColor =0
                    Name ="cmdChart"
                    Caption ="Show Chart"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00000000000000000ad0c010d0c010da ,
                        0x0da0c010a0c010ad0ad0c010d0c010da0da0c010a0c010ad0ad0c010d0c010da ,
                        0x0da0c000a0c010ad0ad0c0dad0c010da0da0c0ada00010ad0ad0c0dadad010da ,
                        0x0da000adada010ad0adadadadad010da0dadadadada010ad0adadadadad000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Chart View of Report"
                    UnicodeAccessKey =104

                    LayoutCachedLeft =8580
                    LayoutCachedTop =6600
                    LayoutCachedWidth =9156
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8580
                            Top =7200
                            Width =480
                            Height =240
                            Name ="Label53"
                            Caption ="C&hart"
                            LayoutCachedLeft =8580
                            LayoutCachedTop =7200
                            LayoutCachedWidth =9060
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =5490
                    Top =2340
                    Width =1770
                    TabIndex =23
                    Name ="cboPriceGroupIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.Description, PriceGroup.id FROM Item INNER JOIN PriceGroup ON "
                        "Item.PriceGroupID=PriceGroup.ID GROUP BY PriceGroup.ID, PriceGroup.Description O"
                        "RDER BY PriceGroup.Description; "
                    ColumnWidths ="1440;1440"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =2580
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4080
                            Top =2340
                            Width =1350
                            Height =240
                            Name ="Label55"
                            Caption ="Price Group"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =2340
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =2580
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =7440
                    Top =2340
                    Width =1770
                    TabIndex =24
                    Name ="cboPriceGroupIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.Description, PriceGroup.id FROM Item INNER JOIN PriceGroup ON "
                        "Item.PriceGroupID=PriceGroup.ID GROUP BY PriceGroup.ID, PriceGroup.Description O"
                        "RDER BY PriceGroup.Description; "
                    ColumnWidths ="1440;1440"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =2340
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =2580
                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =69
                    Left =5820
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =25
                    Name ="cmdMail"
                    Caption ="Command57"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x070707070707070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000ff000000fffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffb000000000000000000000000fffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000000000000000fffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707c2686d136a13000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="EMail Report as an Attachment"
                    UnicodeAccessKey =69

                    LayoutCachedLeft =5820
                    LayoutCachedTop =6600
                    LayoutCachedWidth =6396
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =5865
                            Top =7200
                            Width =435
                            Height =240
                            Name ="Label58"
                            Caption ="&EMail"
                            LayoutCachedLeft =5865
                            LayoutCachedTop =7200
                            LayoutCachedWidth =6300
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5250
                    Top =3840
                    Width =1770
                    TabIndex =26
                    Name ="txtNumericRangeFrom"
                    Format ="Standard"
                    DefaultValue ="0"

                    LayoutCachedLeft =5250
                    LayoutCachedTop =3840
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =4080
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =4080
                            Top =3840
                            Width =1350
                            Height =240
                            Name ="Label60"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =3840
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =4080
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7320
                    Top =3840
                    Width =1770
                    TabIndex =27
                    Name ="txtNumericRangeTo"
                    DefaultValue ="9999"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =3840
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =4080
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =84
                    Left =9240
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =28
                    Name ="cmdPrintChart"
                    Caption ="Print Chart"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad00000000000dadd0888888888080da ,
                        0x000000000000080d0888888bbb88000a088888877788080d0000000000000880 ,
                        0x0888888888808080d000000000080800ad0ffffffff08080dad0f00000f0000a ,
                        0xada0ffffffff0daddada0f00000f0adaadad0ffffffff0addadad000000000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Print Chart View"
                    UnicodeAccessKey =116

                    LayoutCachedLeft =9240
                    LayoutCachedTop =6600
                    LayoutCachedWidth =9816
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9240
                            Top =7200
                            Width =855
                            Height =240
                            Name ="Label63"
                            Caption ="Prin&t Chart"
                            LayoutCachedLeft =9240
                            LayoutCachedTop =7200
                            LayoutCachedWidth =10095
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =88
                    Left =6480
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =29
                    Name ="cmdOutputToFile"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070700000000000000000000000000000000000000 ,
                        0x000000000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706060607060706 ,
                        0x070607000707070707070707070000f806070600000000000000000000000000 ,
                        0x060706000707070707070707000000f807060700ffffffffffffffffffffff00 ,
                        0x070607000707070707070700fffb030000070600ffffffffffffffffffffff00 ,
                        0x060706000707070707070707000000f807060700ffffffffffffffffffffff00 ,
                        0x070607000707070707070707070000f806070600000000000000000000000000 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x0607060007070707070707070700000000000000000000000000000000000000 ,
                        0x0000000007070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Export Report to File"
                    UnicodeAccessKey =120

                    LayoutCachedLeft =6480
                    LayoutCachedTop =6600
                    LayoutCachedWidth =7056
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =6525
                            Top =7200
                            Width =555
                            Height =240
                            Name ="Label66"
                            Caption ="E&xport"
                            LayoutCachedLeft =6525
                            LayoutCachedTop =7200
                            LayoutCachedWidth =7080
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =65
                    Left =7830
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =30
                    Name ="cmdArchiveReport"
                    Caption ="cmdArchiveReport"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8070700000000000000000000000000000000000000000000 ,
                        0x0000000000f8f8f8070007070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700000000000000000000000000000000 ,
                        0x000007070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070000000000000000000000000000000000000000 ,
                        0x000007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8070007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff00070707000707070000000000000000000000000000000000000000000000 ,
                        0x0000000000070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Report to File In Reprint Reports Area (Menu Option A-5)"
                    UnicodeAccessKey =65

                    LayoutCachedLeft =7830
                    LayoutCachedTop =6600
                    LayoutCachedWidth =8406
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =7800
                            Top =7200
                            Width =705
                            Height =240
                            Name ="Label69"
                            Caption ="&Archive"
                            LayoutCachedLeft =7800
                            LayoutCachedTop =7200
                            LayoutCachedWidth =8505
                            LayoutCachedHeight =7440
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =5490
                    Top =2940
                    Width =1770
                    TabIndex =31
                    Name ="cboClubIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Club.Name, Club.ID FROM Club WHERE (((Club.NonLiquor)=False)) ORDER BY Cl"
                        "ub.Name; "
                    ColumnWidths ="2880"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =3180
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4080
                            Top =2940
                            Width =1350
                            Height =240
                            Name ="Label71"
                            Caption ="Clubs"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =2940
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =3180
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =7440
                    Top =2940
                    Width =1770
                    TabIndex =32
                    Name ="cboClubIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Club.Name, Club.ID FROM Club WHERE (((Club.NonLiquor)=False)) ORDER BY Cl"
                        "ub.Name; "
                    ColumnWidths ="2880"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =2940
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =3180
                End
                Begin OptionGroup
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =9360
                    Top =3420
                    Width =2342
                    Height =2998
                    TabIndex =33
                    Name ="framDates"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =3420
                    LayoutCachedWidth =11702
                    LayoutCachedHeight =6418
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =9480
                            Top =3300
                            Width =705
                            Height =240
                            BackColor =-2147483633
                            Name ="Label74"
                            Caption ="Dates"
                            LayoutCachedLeft =9480
                            LayoutCachedTop =3300
                            LayoutCachedWidth =10185
                            LayoutCachedHeight =3540
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9546
                            Top =3658
                            OptionValue =1
                            Name ="Option76"

                            LayoutCachedLeft =9546
                            LayoutCachedTop =3658
                            LayoutCachedWidth =9806
                            LayoutCachedHeight =3898
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9776
                                    Top =3630
                                    Width =525
                                    Height =240
                                    Name ="Label77"
                                    Caption ="Today"
                                    LayoutCachedLeft =9776
                                    LayoutCachedTop =3630
                                    LayoutCachedWidth =10301
                                    LayoutCachedHeight =3870
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9546
                            Top =3988
                            OptionValue =2
                            Name ="Option78"

                            LayoutCachedLeft =9546
                            LayoutCachedTop =3988
                            LayoutCachedWidth =9806
                            LayoutCachedHeight =4228
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9776
                                    Top =3960
                                    Width =810
                                    Height =240
                                    Name ="Label79"
                                    Caption ="Yesterday"
                                    LayoutCachedLeft =9776
                                    LayoutCachedTop =3960
                                    LayoutCachedWidth =10586
                                    LayoutCachedHeight =4200
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9546
                            Top =4318
                            OptionValue =3
                            Name ="Option80"

                            LayoutCachedLeft =9546
                            LayoutCachedTop =4318
                            LayoutCachedWidth =9806
                            LayoutCachedHeight =4558
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9776
                                    Top =4290
                                    Width =810
                                    Height =240
                                    Name ="Label81"
                                    Caption ="This Week"
                                    LayoutCachedLeft =9776
                                    LayoutCachedTop =4290
                                    LayoutCachedWidth =10586
                                    LayoutCachedHeight =4530
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9546
                            Top =4648
                            OptionValue =4
                            Name ="Option82"

                            LayoutCachedLeft =9546
                            LayoutCachedTop =4648
                            LayoutCachedWidth =9806
                            LayoutCachedHeight =4888
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9776
                                    Top =4620
                                    Width =855
                                    Height =240
                                    Name ="Label83"
                                    Caption ="This Month"
                                    LayoutCachedLeft =9776
                                    LayoutCachedTop =4620
                                    LayoutCachedWidth =10631
                                    LayoutCachedHeight =4860
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9540
                            Top =5008
                            OptionValue =5
                            Name ="Option86"

                            LayoutCachedLeft =9540
                            LayoutCachedTop =5008
                            LayoutCachedWidth =9800
                            LayoutCachedHeight =5248
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9770
                                    Top =4980
                                    Width =735
                                    Height =240
                                    Name ="Label87"
                                    Caption ="This Year"
                                    LayoutCachedLeft =9770
                                    LayoutCachedTop =4980
                                    LayoutCachedWidth =10505
                                    LayoutCachedHeight =5220
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9540
                            Top =5368
                            OptionValue =6
                            Name ="Option88"

                            LayoutCachedLeft =9540
                            LayoutCachedTop =5368
                            LayoutCachedWidth =9800
                            LayoutCachedHeight =5608
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9770
                                    Top =5340
                                    Width =870
                                    Height =240
                                    Name ="Label89"
                                    Caption ="Last Week"
                                    LayoutCachedLeft =9770
                                    LayoutCachedTop =5340
                                    LayoutCachedWidth =10640
                                    LayoutCachedHeight =5580
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9540
                            Top =5716
                            OptionValue =7
                            Name ="Option91"

                            LayoutCachedLeft =9540
                            LayoutCachedTop =5716
                            LayoutCachedWidth =9800
                            LayoutCachedHeight =5956
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9770
                                    Top =5688
                                    Width =870
                                    Height =240
                                    Name ="Label92"
                                    Caption ="Last Month"
                                    LayoutCachedLeft =9770
                                    LayoutCachedTop =5688
                                    LayoutCachedWidth =10640
                                    LayoutCachedHeight =5928
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =9540
                            Top =6064
                            OptionValue =8
                            Name ="Option93"

                            LayoutCachedLeft =9540
                            LayoutCachedTop =6064
                            LayoutCachedWidth =9800
                            LayoutCachedHeight =6304
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9770
                                    Top =6036
                                    Width =870
                                    Height =240
                                    Name ="Label94"
                                    Caption ="Last Year"
                                    LayoutCachedLeft =9770
                                    LayoutCachedTop =6036
                                    LayoutCachedWidth =10640
                                    LayoutCachedHeight =6276
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =5490
                    Top =1440
                    Width =1770
                    TabIndex =34
                    BoundColumn =1
                    Name ="cboCategoryIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SubCategory.ID, SubCategory.Description FROM SubCategory ORDER BY SubCate"
                        "gory.Description; "
                    ColumnWidths ="0;1950"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =1440
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1680
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4080
                            Top =1440
                            Width =1350
                            Height =240
                            Name ="Label96"
                            Caption ="Category"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =1440
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2310
                    Left =7440
                    Top =1440
                    Width =1770
                    TabIndex =35
                    BoundColumn =1
                    Name ="cboCategoryIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SubCategory.ID, SubCategory.Description FROM SubCategory ORDER BY SubCate"
                        "gory.Description; "
                    ColumnWidths ="0;1950"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =1680
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4080
                    Top =5100
                    Width =5100
                    Height =1260
                    TabIndex =36
                    BackColor =10092543
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =5100
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =6360
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4440
                    Top =7560
                    Width =3960
                    TabIndex =37
                    BackColor =39423
                    Name ="txtQueryName"
                    ControlSource ="QueryName"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =7560
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =7800
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =3300
                            Top =7560
                            Width =1080
                            Height =240
                            Name ="Label101"
                            Caption ="Query Name:"
                            LayoutCachedLeft =3300
                            LayoutCachedTop =7560
                            LayoutCachedWidth =4380
                            LayoutCachedHeight =7800
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =76
                    Left =7140
                    Top =6600
                    Width =576
                    Height =576
                    TabIndex =38
                    Name ="cmdExportExcel"
                    Caption ="Command104"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Report to File"
                    UnicodeAccessKey =108
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000030000000b00000014000000190000001d ,
                        0x000000230000002a000000320000003800000040000000470000004c0000003b ,
                        0x0000002700000005000000000000001000000020000000400000005000000050 ,
                        0x4040407060686080a0a0a0a0b0b0b0b0c0c8c0c0d0e0e0e0e0f0e0f0d0d8d0c0 ,
                        0x0000004000000020106010c0105810d080a080f0c0d0c0f0d0e0d0ffffffffff ,
                        0xfffffffffffffffff0fffffff0fffffff0fffffff0fffffff0fff0ffc0d8c0f0 ,
                        0x0000004000000020106010f0a0b8a0ffc0d0c0ffffffffffffffffffffffffff ,
                        0xfffffffff0fffffff0ffffff90c0a0ff107030ff107030fff0f8f0ffe0e8e0ff ,
                        0x0000004000000020106810f0b0d0c0ff80a080ffffffffffc0e0c0ff80c080ff ,
                        0x70b870ff70b870ff60b070ff107030ff107030ff60a870ffe0f0e0fff0f8f0ff ,
                        0x0000004000000020106810f0c0d0c0ff407840ffffffffffffffffffc0e0c0ff ,
                        0x70b870ff70b870ff208040ff107030ff60a070ff60b070ffa0d0a0fff0ffffff ,
                        0x4040406000000020107010f0c0d0c0ff206020fff0f0f0ffffffffffffffffff ,
                        0xe0f0e0ff409860ff107030ff207840ffb0e0c0fffffffffff0fffffff0ffffff ,
                        0x7070707000000020207020f0c0d0c0ff206020ffb0c8b0ffffffffffffffffff ,
                        0xc0d8d0ff107030ff107030ff60a860ff70b870ffa0d0a0fff0fff0fff0ffffff ,
                        0x9098909000000020207820f0c0d0c0ff206820ff90a890fffffffffff0f8f0ff ,
                        0x107030ff107030ff90c0a0ffe0f0e0ff70b870ff70b870ff90c890fff0f8f0ff ,
                        0xb0b0b0a000000020207820f0d0e0d0ff70a870ff90b090ffffffffff409060ff ,
                        0x308050ff60a070ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xd0d0d0a000000020208020f0d0e8d0ff80b880ff80b080ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xf0f0f0b000000020208020e0d0e8d0ffa0c8a0ff80b880ffd0e8d0ffd0e0d0ff ,
                        0xd0e0d0ffd0e0d0ffc0d8c0ffb0d8b0ffa0d0a0ff90c8a0ffd0e8d0ff50a050e0 ,
                        0xffffff2000000000208020a0a0c890ffe0f0e0ff90c890ff90c890ff90c890ff ,
                        0x90c890ff90c890ff90c890ff90c890ff90c890ff90c890fff0f8f0ff208020d0 ,
                        0x000000000000000030882030308820ffc0e0c0fff0f8f0ffb0d8b0ffa0d0a0ff ,
                        0xa0d0a0ffa0d0a0ffa0d0a0ffa0d0a0ffa0d0a0ff90d0a0fff0f8f0ff308820d0 ,
                        0x00000000000000000000000030882060308820f090c890ffd0e8d0fff0f8f0ff ,
                        0xf0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0ff308820d0 ,
                        0x0000000000000000000000000000000030902020309020a0309020e0309020ff ,
                        0x309020ff309020ff309020ff309020ff309020ff309020ff309020ff309020d0 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7140
                    LayoutCachedTop =6600
                    LayoutCachedWidth =7716
                    LayoutCachedHeight =7176
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =7140
                            Top =7200
                            Width =555
                            Height =240
                            Name ="Label105"
                            Caption ="Exce&l"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =7200
                            LayoutCachedWidth =7695
                            LayoutCachedHeight =7440
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PrintInventoryReport.cls"
