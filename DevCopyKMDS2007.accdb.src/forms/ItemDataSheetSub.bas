Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9960
    DatasheetFontHeight =10
    ItemSuffix =63
    Left =6150
    Top =810
    Right =16920
    Bottom =3555
    HelpContextId =4200
    Filter ="([Lookup_VendorID].[Name]=\"LUCKY 7\")"
    ShortcutMenuBar ="KMDSPopup"
    OrderBy ="Lookup_SubCategoryID.Description, Item.Description"
    OnFilter ="[Event Procedure]"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x435d42196a94e240
    End
    RecordSource ="Item"
    Caption ="Item"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1620
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =660
                    Height =240
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =780
                    Top =60
                    Width =1620
                    Height =240
                    Name ="SizeID_Label"
                    Caption ="SizeID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2460
                    Top =60
                    Width =1620
                    Height =240
                    Name ="SalesDepartmentID_Label"
                    Caption ="SalesDepartmentID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4140
                    Top =60
                    Width =1620
                    Height =240
                    Name ="VendorID_Label"
                    Caption ="VendorID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5820
                    Top =60
                    Width =1620
                    Height =240
                    Name ="PriceGroupID_Label"
                    Caption ="PriceGroupID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7500
                    Top =60
                    Width =1620
                    Height =240
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =9180
                    Top =60
                    Width =660
                    Height =240
                    Name ="Onhand_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4080
                    Top =840
                    Width =660
                    Height =240
                    Name ="OrderAt_Label"
                    Caption ="OrderAt"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4800
                    Top =840
                    Width =660
                    Height =240
                    Name ="OrderTo_Label"
                    Caption ="OrderTo"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =3960
                    Top =1380
                    Width =1620
                    Height =240
                    Name ="Price_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Top =360
                    Width =1620
                    Height =240
                    Name ="Cost_Label"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1680
                    Top =360
                    Width =1620
                    Height =240
                    Name ="Deposit_Label"
                    Caption ="Deposit"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3420
                    Top =360
                    Width =720
                    Height =240
                    Name ="UnitsPerCase_Label"
                    Caption ="UnitsPerCase"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4200
                    Top =360
                    Width =780
                    Height =240
                    Name ="UnitsPerOrder_Label"
                    Caption ="UnitsPerOrder"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5040
                    Top =360
                    Width =780
                    Height =240
                    Name ="UnitsPerPack_Label"
                    Caption ="UnitsPerPack"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =5880
                    Top =360
                    Width =660
                    Height =240
                    Name ="SubPack_Label"
                    Caption ="SubPack"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =6900
                    Top =360
                    Width =660
                    Height =240
                    Name ="Location_Label"
                    Caption ="Location"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =7620
                    Top =360
                    Width =660
                    Height =240
                    Name ="QtyPack_Label"
                    Caption ="QtyPack"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =8340
                    Top =360
                    Width =1620
                    Height =240
                    Name ="QtyPackPrice_Label"
                    Caption ="QtyPackPrice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =4200
                    Top =1140
                    Width =1080
                    Height =240
                    Name ="StockNumber_Label"
                    Caption ="StockNumber"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =5340
                    Top =1140
                    Width =480
                    Height =240
                    Name ="PLUList_Label"
                    Caption ="PLUList"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =6360
                    Top =600
                    Width =720
                    Height =240
                    Name ="DateLastSold_Label"
                    Caption ="DateLastSold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =7140
                    Top =600
                    Width =600
                    Height =240
                    Name ="SetupDate_Label"
                    Caption ="SetupDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =7800
                    Top =600
                    Width =975
                    Height =240
                    Name ="ChangeDate_Label"
                    Caption ="ChangeDate"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =1995
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =600
                    ColumnOrder =1
                    BackColor =10079487
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ColumnCount =2
                    ListRows =20
                    Left =780
                    Top =60
                    Width =1620
                    Height =255
                    ColumnWidth =1245
                    ColumnOrder =12
                    TabIndex =1
                    Name ="SizeID"
                    ControlSource ="SizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.ID, Size.Description FROM [Size] ORDER BY Size.Description; "
                    ColumnWidths ="0;1440"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    ListRows =20
                    Left =2460
                    Top =60
                    Width =1620
                    Height =255
                    ColumnWidth =1860
                    ColumnOrder =0
                    TabIndex =2
                    BoundColumn =1
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.Description, SalesDepartment.ID FROM SalesDepartment ORDE"
                        "R BY SalesDepartment.Description; "
                    ColumnWidths ="2160;720"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ColumnCount =2
                    ListRows =20
                    Left =4140
                    Top =60
                    Width =1620
                    Height =255
                    ColumnWidth =1680
                    ColumnOrder =3
                    TabIndex =3
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name FROM Vendor ORDER BY Vendor.Name; "
                    ColumnWidths ="0;2880"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ColumnCount =2
                    ListRows =20
                    Left =5820
                    Top =60
                    Width =1620
                    Height =255
                    ColumnWidth =1335
                    ColumnOrder =14
                    TabIndex =4
                    Name ="PriceGroupID"
                    ControlSource ="PriceGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.ID, PriceGroup.Description FROM PriceGroup ORDER BY PriceGroup"
                        ".Description; "
                    ColumnWidths ="0;1440"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =7500
                    Top =60
                    Width =1620
                    Height =255
                    ColumnWidth =2280
                    ColumnOrder =2
                    TabIndex =5
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =9180
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =855
                    ColumnOrder =7
                    TabIndex =6
                    Name ="Onhand"
                    ControlSource ="Onhand"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    Left =2820
                    Top =1020
                    Width =660
                    Height =255
                    ColumnWidth =855
                    ColumnOrder =5
                    TabIndex =7
                    Name ="OrderAt"
                    ControlSource ="OrderAt"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    Left =3540
                    Top =1020
                    Width =660
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =6
                    TabIndex =8
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Top =420
                    Width =1620
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =9
                    TabIndex =9
                    Name ="Price"
                    ControlSource ="Price"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =1680
                    Top =420
                    Width =1620
                    Height =255
                    ColumnWidth =765
                    ColumnOrder =8
                    TabIndex =10
                    Name ="Cost"
                    ControlSource ="Cost"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3360
                    Top =420
                    Width =1620
                    Height =255
                    ColumnWidth =840
                    ColumnOrder =15
                    TabIndex =11
                    Name ="Deposit"
                    ControlSource ="Deposit"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =5100
                    Top =420
                    Width =720
                    Height =255
                    ColumnWidth =1380
                    ColumnOrder =17
                    TabIndex =12
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =5880
                    Top =420
                    Width =780
                    Height =255
                    ColumnWidth =1410
                    ColumnOrder =4
                    TabIndex =13
                    Name ="UnitsPerOrder"
                    ControlSource ="UnitsPerOrder"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =6720
                    Top =420
                    Width =780
                    Height =255
                    ColumnWidth =1380
                    ColumnOrder =18
                    TabIndex =14
                    Name ="UnitsPerPack"
                    ControlSource ="UnitsPerPack"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =7560
                    Top =420
                    Width =960
                    Height =255
                    ColumnWidth =1230
                    ColumnOrder =19
                    TabIndex =15
                    Name ="SubPack"
                    ControlSource ="SubPack"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8580
                    Top =420
                    Width =660
                    Height =255
                    ColumnWidth =915
                    ColumnOrder =20
                    TabIndex =16
                    Name ="Location"
                    ControlSource ="Location"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =9300
                    Top =420
                    Width =660
                    Height =255
                    ColumnWidth =945
                    ColumnOrder =24
                    TabIndex =17
                    Name ="QtyPack"
                    ControlSource ="QtyPack"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =2940
                    Top =1380
                    Width =1620
                    Height =255
                    ColumnWidth =1395
                    ColumnOrder =23
                    TabIndex =18
                    Name ="QtyPackPrice"
                    ControlSource ="QtyPackPrice"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =6300
                    Top =720
                    Width =1080
                    Height =255
                    ColumnWidth =1365
                    ColumnOrder =16
                    TabIndex =19
                    Name ="StockNumber"
                    ControlSource ="StockNumber"

                End
                Begin CheckBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =7440
                    Top =720
                    ColumnWidth =870
                    ColumnOrder =11
                    TabIndex =20
                    Name ="PLUList"
                    ControlSource ="PLUList"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =7980
                    Top =720
                    Width =720
                    Height =255
                    ColumnWidth =1350
                    ColumnOrder =13
                    TabIndex =21
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =2940
                    Top =1740
                    Width =600
                    Height =255
                    ColumnWidth =1140
                    ColumnOrder =21
                    TabIndex =22
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =3600
                    Top =1740
                    Width =975
                    Height =255
                    ColumnWidth =1245
                    ColumnOrder =22
                    TabIndex =23
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3555
                    Left =1770
                    Top =720
                    Width =3840
                    Height =300
                    ColumnWidth =960
                    ColumnOrder =10
                    TabIndex =24
                    BackColor =16777215
                    ForeColor =0
                    Name ="SubCategoryID"
                    ControlSource ="SubCategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SubCategory.ID, SubCategory.Description FROM SubCategory ORDER BY [Descri"
                        "ption]; "
                    ColumnWidths ="0;3554"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =720
                            Width =1500
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="Category_Label"
                            Caption ="Category"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =600
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =120
                    BackColor =16777215
                    ForeColor =0
                    Name ="txtTotalOnhand"
                    ControlSource ="=Sum([onhand])"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =120
                            Width =1050
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label56"
                            Caption ="Total Onhand"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4440
                    Top =60
                    TabIndex =1
                    BackColor =16777215
                    ForeColor =0
                    Name ="txtTotalRetail"
                    ControlSource ="=Sum([onhand]*[Price])"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3000
                            Top =60
                            Width =1050
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label59"
                            Caption ="Total Retail"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7560
                    Top =60
                    TabIndex =2
                    BackColor =16777215
                    ForeColor =0
                    Name ="txtTotalCost"
                    ControlSource ="=Sum([onhand]*[Cost])"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6120
                            Top =60
                            Width =1050
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label61"
                            Caption ="Total Retail"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ItemDataSheetSub.cls"
