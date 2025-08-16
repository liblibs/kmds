Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    MaxButton = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5160
    DatasheetFontHeight =10
    ItemSuffix =96
    Right =5055
    Bottom =8580
    HelpContextId =4200
    OrderBy ="myDescription"
    RecSrcDt = Begin
        0xd0d2da828998e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="DataSheetTemplate"
    Caption ="Criteria"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    MenuBar ="KMDS"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =600
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2100
                    Top =300
                    Width =2520
                    Height =300
                    ColumnOrder =0
                    Name ="myDescription"
                    ControlSource ="myDescription"
                    FontName ="Trebuchet MS"
                    Tag ="myDescription"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2100
                            Top =60
                            Width =945
                            Height =225
                            Name ="Label88"
                            Caption ="Description"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =25
                    ListWidth =2880
                    Left =60
                    Top =300
                    Width =1980
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="cmbmyDescription"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DataSheetTemplate.myID, DataSheetTemplate.myDescription FROM DataSheetTem"
                        "plate ORDER BY DataSheetTemplate.myDescription; "
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =60
                            Top =60
                            Width =900
                            Height =245
                            Name ="Description_Label"
                            Caption ="Lookup"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =6960
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =3780
                    Height =180
                    Name ="chkVendorID"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Items are purchased from vendors for resale."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =187
                            Top =3780
                            Width =1290
                            Height =240
                            Name ="Label1"
                            Caption ="Vendor/Supplier"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =1860
                    Height =180
                    TabIndex =1
                    Name ="chkOnhand"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Quantity of items in stock."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =832
                            Top =1860
                            Width =645
                            Height =240
                            Name ="Label4"
                            Caption ="Onhand"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =2700
                    Height =180
                    TabIndex =2
                    Name ="chkOrderAt"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Order point when creating a purchase order.  If the onhand is less then this amo"
                        "unt then create purchase order line."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =742
                            Top =2700
                            Width =735
                            Height =240
                            Name ="Label6"
                            Caption ="Order at"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =2964
                    Height =180
                    TabIndex =3
                    Name ="chkOrderTo"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Used in creating purchase order.  This is the maximum to order to.. Order quanti"
                        "ty is the difference between Order To and the Onhand amount."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =742
                            Top =2970
                            Width =735
                            Height =240
                            Name ="Label8"
                            Caption ="Order to"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =1620
                    Height =180
                    TabIndex =4
                    Name ="chkSizeID"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="This is the container size of the item sold."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =427
                            Top =1620
                            Width =1050
                            Height =240
                            Name ="Label11"
                            Caption ="Package Size"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =600
                    Width =185
                    Height =180
                    TabIndex =5
                    Name ="chkSalesDepartmentID"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Items are grouped by Sales departments"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =97
                            Top =600
                            Width =1380
                            Height =240
                            Name ="Label13"
                            Caption ="Sales Department"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =2400
                    Height =180
                    TabIndex =6
                    Name ="chkPrice"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Selling price, the customer pays this price."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =547
                            Top =2407
                            Width =930
                            Height =240
                            Name ="Label15"
                            Caption ="Retail Price"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =2107
                    Height =180
                    TabIndex =7
                    Name ="chkCost"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Amount paid to the vendor to purchase this item for resale."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =322
                            Top =2100
                            Width =1155
                            Height =240
                            Name ="Label17"
                            Caption ="Purchase Cost"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =3486
                    Height =180
                    TabIndex =8
                    Name ="chkDeposit"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Amount of bottle deposit.  When an item is sold with a deposit amount then a  se"
                        "cond line is added to the sale for the deposit amount."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =832
                            Top =3486
                            Width =645
                            Height =240
                            Name ="Label19"
                            Caption ="Deposit"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =4320
                    Height =180
                    TabIndex =9
                    Name ="chkUnitsPerCase"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Quantity of items in 1 case."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =517
                            Top =4320
                            Width =960
                            Height =240
                            Name ="Label21"
                            Caption ="Units /Case"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =4560
                    Height =180
                    TabIndex =10
                    Name ="chkUnitsPerOrder"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Items are ordered in multiples of this amount.  If  the vendor requires to order"
                        " this item  in case quantities then the Units Per Order  must be at least the Un"
                        "its per Case quantity."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =472
                            Top =4560
                            Width =1005
                            Height =240
                            Name ="Label24"
                            Caption ="Units/Order"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =4800
                    Height =180
                    TabIndex =11
                    Name ="chkUnitsPerPack"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="This the amount of packs in a carton; generally this is 10 for 10 packs in a car"
                        "ton of cigarettes.  Used when transferring cartons to packs."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =562
                            Top =4800
                            Width =915
                            Height =240
                            Name ="Label26"
                            Caption ="Units/Pack"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =3240
                    Height =180
                    TabIndex =12
                    Name ="chkPriceGroupID"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Allows grouping of similarly price items.  If  the selling price of an item is c"
                        "hanged and it is in a Price Group then all the other items within the price grou"
                        "p are also changed."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =517
                            Top =3240
                            Width =960
                            Height =240
                            Name ="Label28"
                            Caption ="Price Group"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =5042
                    Height =180
                    TabIndex =13
                    Name ="chkSubPack"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Used when transferring cartons to packs.  This is the code for the pack. This is"
                        " how the carton and pack are linked.  The pack's item code is kept in this field"
                        "."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =787
                            Top =5042
                            Width =690
                            Height =240
                            Name ="Label30"
                            Caption ="SubPack"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =5280
                    Height =180
                    TabIndex =14
                    Name ="chkLocation"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Location this item is found in the store.  See manual for a suggestion on how to"
                        " implement a system for locations."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =352
                            Top =5280
                            Width =1125
                            Height =240
                            Name ="Label32"
                            Caption ="Shelf Location"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =5580
                    Height =180
                    TabIndex =15
                    Name ="chkQtyPack"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Used by the Cash Register to give Quantity Discounts for an item.  For example: "
                        "If Pepsi is $0.50 each or 6 @ $1.99 then 6 is the Quantity amount."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =727
                            Top =5580
                            Width =750
                            Height =240
                            Name ="Label34"
                            Caption ="Quantity"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =4020
                    Height =180
                    TabIndex =16
                    Name ="chkStockNumber"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="This is the item number the vendor uses to identify an item.  This is printed on"
                        " the purchase orders."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =472
                            Top =4020
                            Width =1005
                            Height =240
                            Name ="Label36"
                            Caption ="Vendors SKU"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =5820
                    Height =180
                    TabIndex =17
                    Name ="chkPLUList"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Used by the Cash Register to keep a short list of items they do not have barcode"
                        "s and must be entered by their code number."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =457
                            Top =5820
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="POs PLU List"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =6060
                    Height =180
                    TabIndex =18
                    Name ="chkDateLastSold"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Date this item was last sold."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =712
                            Top =6060
                            Width =765
                            Height =240
                            Name ="Label40"
                            Caption ="Date Sold"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =6300
                    TabIndex =19
                    Name ="chkSetupDate"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Date this item was initally setup."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =982
                            Top =6300
                            Width =495
                            Height =240
                            Name ="Label42"
                            Caption ="Setup"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1560
                    Top =6540
                    TabIndex =20
                    Name ="chkChangeDate"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Date this item was last changed."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =757
                            Top =6540
                            Width =720
                            Height =240
                            Name ="Label44"
                            Caption ="Changed"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =1620
                    Width =375
                    TabIndex =22
                    Name ="txtSizeID"
                    ControlSource ="SizeID"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =600
                    Width =375
                    TabIndex =23
                    Name ="txtSalesDpt"
                    ControlSource ="SalesDepartmentID"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =3780
                    Width =375
                    TabIndex =24
                    Name ="txtVendor"
                    ControlSource ="VendorID"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =3240
                    Width =375
                    TabIndex =25
                    Name ="txtPriceGroup"
                    ControlSource ="PriceGroupID"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =1860
                    Width =375
                    TabIndex =26
                    Name ="txtOnhand"
                    ControlSource ="Onhand"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =2705
                    Width =375
                    TabIndex =27
                    Name ="txtOrderAt"
                    ControlSource ="OrderAt"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =2940
                    Width =375
                    TabIndex =28
                    Name ="txtOrderTo"
                    ControlSource ="OrderTo"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =2422
                    Width =375
                    TabIndex =29
                    Name ="txtPrice"
                    ControlSource ="Price"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =2115
                    Width =375
                    TabIndex =30
                    Name ="txtCost"
                    ControlSource ="Cost"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =3480
                    Width =375
                    TabIndex =31
                    Name ="txtDeposit"
                    ControlSource ="Deposit"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =4320
                    Width =375
                    TabIndex =32
                    Name ="txtUnitsPerCase"
                    ControlSource ="UnitsPerCase"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =4560
                    Width =375
                    TabIndex =33
                    Name ="txtUnitsPerOrder"
                    ControlSource ="UnitsPerOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =4800
                    Width =375
                    TabIndex =34
                    Name ="txtUnitsPerPack"
                    ControlSource ="UnitsPerPack"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =5040
                    Width =375
                    TabIndex =35
                    Name ="txtSubPack"
                    ControlSource ="SubPack"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =5280
                    Width =375
                    TabIndex =36
                    Name ="txtLocation"
                    ControlSource ="Location"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =5580
                    Width =375
                    TabIndex =37
                    Name ="txtQtyPack"
                    ControlSource ="QtyPack"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =4020
                    Width =375
                    TabIndex =38
                    Name ="txtStockNumber"
                    ControlSource ="StockNumber"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =5820
                    Width =375
                    TabIndex =39
                    Name ="txtPLUList"
                    ControlSource ="QtyPackPrice"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =6060
                    Width =375
                    TabIndex =40
                    Name ="txtDateLastSold"
                    ControlSource ="DateLastSold"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =6300
                    Width =375
                    TabIndex =41
                    Name ="txtSetupDate"
                    ControlSource ="SetupDate"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =6540
                    Width =375
                    TabIndex =42
                    Name ="txtChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =82
                    Left =300
                    Top =180
                    Width =840
                    TabIndex =44
                    Name ="cmdClearAll"
                    Caption ="Clea&r All"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =114

                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =1620
                    Width =375
                    TabIndex =45
                    Name ="txtSizeIDSB"
                    ControlSource ="SizeIDSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Lookup_SizeID.Description"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =600
                    Width =375
                    TabIndex =46
                    Name ="txtSalesDptSB"
                    ControlSource ="IDSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Lookup_SalesDepartmentID.Description"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =3780
                    Width =375
                    TabIndex =47
                    Name ="txtVendorSB"
                    ControlSource ="VendorIDSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Lookup_VendorID.Name"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =3240
                    Width =375
                    TabIndex =48
                    Name ="txtPriceGroupSB"
                    ControlSource ="PriceGroupIDSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Lookup_PriceGroupID.Description"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =1860
                    Width =375
                    TabIndex =49
                    Name ="txtOnhandSB"
                    ControlSource ="OnhandSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Onhand"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =2705
                    Width =375
                    TabIndex =50
                    Name ="txtOrderAtSB"
                    ControlSource ="OrderAtSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="OrderAt"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =2940
                    Width =375
                    TabIndex =51
                    Name ="txtOrderToSB"
                    ControlSource ="OrderToSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="OrderTo"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =2422
                    Width =375
                    TabIndex =52
                    Name ="txtPriceSB"
                    ControlSource ="PriceSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Price"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =2115
                    Width =375
                    TabIndex =53
                    Name ="txtCostSB"
                    ControlSource ="CostSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Cost"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =3480
                    Width =375
                    TabIndex =54
                    Name ="txtDepositSB"
                    ControlSource ="DepositSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Deposit"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =4320
                    Width =375
                    TabIndex =55
                    Name ="txtUnitsPerCaseSB"
                    ControlSource ="UnitsPerCaseSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="UnitsPerCase"

                End
                Begin TextBox
                    IMEHold = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =4560
                    Width =375
                    TabIndex =56
                    Name ="txtUnitsPerOrderSB"
                    ControlSource ="UnitsPerOrderSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="UnitsPerOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =4800
                    Width =375
                    TabIndex =57
                    Name ="txtUnitsPerPackSB"
                    ControlSource ="UnitsPerPackSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="UnitsPerPack"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =5040
                    Width =375
                    TabIndex =58
                    Name ="txtSubPackSB"
                    ControlSource ="SubPackSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="SubPack"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =5280
                    Width =375
                    TabIndex =59
                    Name ="txtLocationSB"
                    ControlSource ="LocationSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Location"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =5580
                    Width =375
                    TabIndex =60
                    Name ="txtQtyPackSB"
                    ControlSource ="QtyPackSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="QtyPack"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =4020
                    Width =375
                    TabIndex =61
                    Name ="txtStockNumberSB"
                    ControlSource ="StockNumberSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="StockNumber"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =5820
                    Width =375
                    TabIndex =62
                    Name ="txtPLUListSB"
                    ControlSource ="PLUListSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="PLUList"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =6060
                    Width =375
                    TabIndex =63
                    Name ="txtDateLastSoldSB"
                    ControlSource ="DateLastSoldSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="DateLastSold"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =6300
                    Width =375
                    TabIndex =64
                    Name ="txtSetupDateSB"
                    ControlSource ="SetupDateSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="SetupDate"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =6540
                    Width =375
                    TabIndex =65
                    Name ="txtChangeDateSB"
                    ControlSource ="ChangeDateSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ChangeDate"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =840
                    TabIndex =21
                    Name ="chkSubCategoryID"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Allows Sub-Grouping of Sales Departments; \"Whiskey Straight Bourbon\" is a Cate"
                        "gory of the sales department \"Liquor\""

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =660
                            Top =840
                            Width =780
                            Height =240
                            Name ="Label46"
                            Caption ="Category"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =840
                    Width =375
                    TabIndex =43
                    Name ="txtSubCategoryID"
                    ControlSource ="SubCategoryID"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =840
                    Width =375
                    TabIndex =66
                    Name ="txtSubCategoryIDSB"
                    ControlSource ="SubCategoryIDSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Lookup_SubCategoryID.Description"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =1080
                    Height =180
                    TabIndex =67
                    Name ="chkID"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    Tag ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1169
                            Top =1080
                            Width =240
                            Height =240
                            Name ="Label76"
                            Caption ="ID"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =1080
                    Width =375
                    TabIndex =68
                    Name ="txtID"
                    ControlSource ="ID"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =1080
                    Width =375
                    TabIndex =69
                    Name ="txtIDSB"
                    ControlSource ="IDSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Item.ID"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1523
                    Top =1320
                    Height =180
                    TabIndex =70
                    Name ="chkDescription"
                    DefaultValue ="No"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="This prints on the cash register receipt when the item  is sold."

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =532
                            Top =1320
                            Width =945
                            Height =240
                            Name ="Label80"
                            Caption ="Description"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1868
                    Top =1320
                    Width =375
                    TabIndex =71
                    Name ="txtDescription"
                    ControlSource ="Description"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="ColumnOrder"

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2438
                    Top =1320
                    Width =375
                    TabIndex =72
                    Name ="txtDescriptionSB"
                    ControlSource ="DescriptionSortOrder"
                    Format ="0;(0);\"\";\"Null\""
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="Item.Description"

                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =120
                    Width =4560
                    Height =6780
                    Name ="boxScreen"
                End
                Begin Label
                    OverlapFlags =247
                    Left =1260
                    Top =240
                    Width =1035
                    Height =240
                    Name ="Label85"
                    Caption ="Show / Order"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =2520
                    Top =240
                    Width =375
                    Height =240
                    Name ="Label86"
                    Caption ="Sort"
                    FontName ="Trebuchet MS"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =68
                    Left =3060
                    Top =2820
                    Width =666
                    Height =606
                    TabIndex =73
                    ForeColor =0
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =68

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =3735
                            Top =3000
                            Width =630
                            Height =210
                            Name ="lblDeleteRecord"
                            Caption ="&Delete"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =83
                    Left =3060
                    Top =2160
                    Width =666
                    Height =606
                    TabIndex =74
                    ForeColor =0
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
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
                        0x0000000000070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =83

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =3735
                            Top =2340
                            Width =555
                            Height =210
                            Name ="lblSaveRecord"
                            Caption ="&Save"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =85
                    Left =3060
                    Top =1500
                    Width =666
                    Height =606
                    TabIndex =75
                    ForeColor =0
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
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
                        0x0707070707070707070707070707070707070707070707070707070707070700 ,
                        0xfb01fb01000707070707070707070707070707070707070707070707070700fb ,
                        0x01fb01000107070707070707070707070707070707070707070707070700fb01 ,
                        0xfb010001fb070707070707070707070707070707070707070707070700fb01fb ,
                        0x010001fb010707070707070707070707070707070707070707070700fb01fb01 ,
                        0x0001fb01fb07070707070707070707070707070707070707070700fb01fb0100 ,
                        0x01fb01fb00070707070707070707070707070707070707070700fb01fb010001 ,
                        0xfb01fb00fb0707070707070707070707070707070707070700fb01fb010001fb ,
                        0x01fb00fb0007070707070707070707070707070707070700000000010001fb01 ,
                        0xfb00fb0007070707070707070707070707070707070700020202000001fb01fb ,
                        0x00fb00070707070707070707070707070707070707000000000202020001fb00 ,
                        0xfb0007070707070707070707070707070707070700fb01fb00000202000000fb ,
                        0x0007070707070707070707070707070707070700fb01fb01fb0100020200fb00 ,
                        0x07070707070707070707070707070707070700fb01fb01fb01fb000002000007 ,
                        0x07070707070707070707070707070707070000000001fb01fb01fb0002000707 ,
                        0x0707070707070707070707070707070700020202000001fb01fb010000070707 ,
                        0x070707070707070707070707070707000000000202020001fb01fb0007070707 ,
                        0x07070707070707070707070707070007fd0700000202000001fb000707070707 ,
                        0x070707070707070707070707070007fd07fd07fd00020200fb00070707070707 ,
                        0x070707070707070707070707000000000007fd07000002000007070707070707 ,
                        0x07070707070707070707070700fdfffdff0007fd070002000707070707070707 ,
                        0x070707070707070707070700fdfffdfffd00fd07fd0000070707070707070707 ,
                        0x070707070707070707070700fffdfffdff0000fd070007070707070707070707 ,
                        0x070707070707070707070700fdfffdfffdff0007000707070707070707070707 ,
                        0x070707070707070707070700fffdfffdfffd0000070707070707070707070707 ,
                        0x07070707070707070707070700fffdfffdff0007070707070707070707070707 ,
                        0x07070707070707070707070700fdfffdff000707070707070707070707070707 ,
                        0x0707070707070707070707070700000000070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="Trebuchet MS"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Changes Made to This Record"
                    UnicodeAccessKey =85

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =3735
                            Top =1620
                            Width =585
                            Height =255
                            Name ="lblUndo"
                            Caption ="&Undo"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =2940
                    Top =720
                    Width =1560
                    Height =3660
                    Name ="Box96"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =67
                    Left =3060
                    Top =3540
                    Width =666
                    Height =606
                    TabIndex =76
                    ForeColor =0
                    Name ="cmdfrmClose"
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =3772
                            Top =3540
                            Width =510
                            Height =525
                            Name ="Label65"
                            Caption ="&Close Form"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =65
                    Left =3060
                    Top =840
                    Width =666
                    Height =606
                    TabIndex =77
                    ForeColor =0
                    Name ="cmdfrmAddNew"
                    Caption ="Command67"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddcdadadadadadadaaccdadadada4adaddcccdadad4d4d4da ,
                        0xaccccdadad444daddcccccda4444444aaccccdadad444daddcccdadad4d4d4da ,
                        0xaccdadadada4adaddcdadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add a New Record"
                    UnicodeAccessKey =65

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =3735
                            Top =1020
                            Width =720
                            Height =240
                            Name ="Label68"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1020
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =600
                    Width =3900
                    Height =1020
                    Name ="txtFilter"
                    ControlSource ="Filter"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    Tag ="Filter"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Width =480
                            Height =240
                            Name ="Label89"
                            Caption ="Filter:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ItemDataSheetSelect.cls"
