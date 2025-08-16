Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13695
    DatasheetFontHeight =10
    ItemSuffix =138
    Top =540
    Right =13680
    Bottom =5040
    OnUnload ="[Event Procedure]"
    Filter ="[ID]=6208"
    ShortcutMenuBar ="KMDSPopup"
    RecSrcDt = Begin
        0xa2eacdec3a95e240
    End
    RecordSource ="PurchaseOrder"
    Caption ="Purchase Orders KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software"
        " ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Sof"
        "tware ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMD"
        "S Software ©  KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnResize ="[Event Procedure]"
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
            Visible = NotDefault
            CanGrow = NotDefault
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =5820
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    Left =180
                    Top =720
                    Width =13470
                    Height =2580
                    TabIndex =32
                    Name ="PurchaseOrdersOnFileChartSalesTotalSub"
                    SourceObject ="Form.PurchaseOrdersOnFileChartSalesTotal"

                    LayoutCachedLeft =180
                    LayoutCachedTop =720
                    LayoutCachedWidth =13650
                    LayoutCachedHeight =3300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =4440
                    Top =345
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalPriceRepacks"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumRepackDollars"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =345
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =4440
                            Top =105
                            Width =2055
                            Height =255
                            BackColor =13750737
                            Name ="Label104"
                            Caption ="Repacks"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =105
                            LayoutCachedWidth =6495
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =6615
                    Top =345
                    Width =1185
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalPriceCases"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumCaseDollars"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =6615
                    LayoutCachedTop =345
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =6600
                            Top =105
                            Width =2055
                            Height =255
                            BackColor =13750737
                            Name ="Label107"
                            Caption ="Full Cases"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6600
                            LayoutCachedTop =105
                            LayoutCachedWidth =8655
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =11
                    ListWidth =1920
                    Left =300
                    Top =1920
                    Width =3360
                    TabIndex =1
                    BackColor =8454143
                    Name ="cmbVendorID"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name, Vendor.Contact, Vendor.Address1, Vendor.Address2,"
                        " Vendor.City, Vendor.St, Vendor.Zip, Vendor.Phone1, Vendor.POReportForm, Vendor."
                        "PrintUpcOnPo FROM Vendor; "
                    ColumnWidths ="0;1919"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =180
                            Top =1620
                            Width =1245
                            Height =255
                            Name ="Label74"
                            Caption ="Purchased from"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    Left =300
                    Top =2160
                    Width =3300
                    Height =255
                    TabIndex =2
                    BackColor =8454143
                    Name ="txtContact"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =300
                    Top =2400
                    Width =3300
                    Height =255
                    TabIndex =3
                    BackColor =8454143
                    Name ="txtAdd1"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =300
                    Top =2640
                    Width =3300
                    Height =255
                    TabIndex =4
                    BackColor =8454143
                    Name ="txtAdd2"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =300
                    Top =2880
                    Width =3300
                    Height =255
                    TabIndex =5
                    BackColor =8454143
                    Name ="txtAdd3"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    Left =3840
                    Top =2640
                    Width =3300
                    Height =255
                    TabIndex =6
                    BackColor =8454143
                    Name ="txtPhone1"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3480
                    Top =345
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalQ"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumQuantity"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =345
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =3480
                            Top =105
                            Width =840
                            Height =240
                            BackColor =13750737
                            Name ="Label41"
                            Caption ="Quantity"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =3480
                            LayoutCachedTop =105
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1440
                    Top =360
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalCost"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumCost"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =360
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =1440
                            Top =105
                            Width =1200
                            Height =240
                            BackColor =13750737
                            Name ="Label43"
                            Caption ="Cost"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =1440
                            LayoutCachedTop =105
                            LayoutCachedWidth =2640
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =1620
                    Width =3540
                    Height =1620
                    Name ="boxPurchasedFrom"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =5010
                    Top =2040
                    Width =2190
                    Height =255
                    TabIndex =9
                    BackColor =8454143
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4140
                            Top =2040
                            Width =825
                            Height =240
                            Name ="Order_Label"
                            Caption ="Order No."
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =5010
                    Top =2280
                    Width =2190
                    Height =255
                    TabIndex =10
                    BackColor =8454143
                    Name ="OrdDate"
                    ControlSource ="OrderDate"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =3
                            Left =4050
                            Top =2280
                            Width =915
                            Height =240
                            Name ="OrdDate_Label"
                            Caption ="Order Date"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3720
                    Top =1620
                    Width =3540
                    Height =1620
                    Name ="Box35"
                End
                Begin Label
                    OverlapFlags =255
                    Left =3720
                    Top =1620
                    Width =870
                    Height =285
                    Name ="Label36"
                    Caption ="Order Info"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =9780
                    Top =45
                    Width =3060
                    Height =480
                    FontSize =20
                    FontWeight =700
                    BackColor =32768
                    ForeColor =32768
                    Name ="Label58"
                    Caption ="Purchase Order"
                    FontName ="Trebuchet MS"
                    LayoutCachedLeft =9780
                    LayoutCachedTop =45
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =525
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =60
                    Width =13560
                    Height =5700
                    Name ="BoxBody"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =240
                    Top =360
                    Width =1200
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalPrice"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumPrice"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =240
                    LayoutCachedTop =360
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =2
                            Left =240
                            Top =105
                            Width =1200
                            Height =240
                            BackColor =13750737
                            Name ="Label79"
                            Caption ="Price"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =240
                            LayoutCachedTop =105
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =76
                    Left =8700
                    Top =75
                    Width =185
                    TabIndex =12
                    Name ="Check81"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =76

                    LayoutCachedLeft =8700
                    LayoutCachedTop =75
                    LayoutCachedWidth =8885
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            SpecialEffect =2
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =1
                            Left =8925
                            Top =45
                            Width =795
                            Height =240
                            BackColor =65280
                            ForeColor =-2147483640
                            Name ="Label82"
                            Caption ="&Lines only"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =8925
                            LayoutCachedTop =45
                            LayoutCachedWidth =9720
                            LayoutCachedHeight =285
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =3360
                    Width =13560
                    Height =2460
                    Name ="boxSubFormNormal"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =2700
                    Top =345
                    Width =780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalGP"
                    ControlSource ="=IIf(PurchaseOrdersOnFileSub.Form!txtSumPrice<>0,(PurchaseOrdersOnFileSub.Form!t"
                        "xtSumPrice-PurchaseOrdersOnFileSub.Form!txtSumCost)/PurchaseOrdersOnFileSub.Form"
                        "!txtSumPrice,0)"
                    Format ="Percent"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =345
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =645
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =2685
                            Top =105
                            Width =795
                            Height =255
                            BackColor =13750737
                            Name ="Label97"
                            Caption ="GP"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =2685
                            LayoutCachedTop =105
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5640
                    Top =345
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalQRepacks"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumRepackQty"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =5640
                    LayoutCachedTop =345
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =645
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =7800
                    Top =345
                    Width =855
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalQCases"
                    ControlSource ="=PurchaseOrdersOnFileSub.Form!txtSumCaseQty"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =345
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =645
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =7560
                    Top =1650
                    TabIndex =18
                    Name ="chkSupplyAllItems"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =7785
                            Top =1620
                            Width =2370
                            Height =240
                            Name ="Label109"
                            Caption ="This vendor can supply any Item"
                        End
                    End
                End
                Begin OptionGroup
                    TabStop = NotDefault
                    BackStyle =1
                    OverlapFlags =255
                    Left =180
                    Top =720
                    Width =10860
                    Height =840
                    TabIndex =21
                    BackColor =13434828
                    Name ="optSortBy"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =780
                            Top =1020
                            Width =570
                            Height =240
                            Name ="Label52"
                            Caption ="Sort by"
                            FontName ="Trebuchet MS"
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =1635
                            Top =1080
                            Width =245
                            OptionValue =1
                            Name ="Option54"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =1440
                                    Top =780
                                    Width =585
                                    Height =285
                                    Name ="Label55"
                                    Caption ="Line Id"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =2421
                            Top =1080
                            Width =245
                            OptionValue =4
                            Name ="Option60"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =2100
                                    Top =780
                                    Width =750
                                    Height =285
                                    Name ="Label61"
                                    Caption ="Quantity"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =3195
                            Top =1080
                            Width =245
                            OptionValue =2
                            Name ="Option56"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =2940
                                    Top =780
                                    Width =630
                                    Height =285
                                    Name ="Label57"
                                    Caption ="Item ID"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =3975
                            Top =1080
                            Width =245
                            OptionValue =3
                            Name ="Option58"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =3660
                                    Top =780
                                    Width =945
                                    Height =285
                                    Name ="Label59"
                                    Caption ="Description"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =4995
                            Top =1080
                            Width =245
                            OptionValue =13
                            Name ="Option129"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextAlign =2
                                    Left =4710
                                    Top =780
                                    Width =735
                                    Height =285
                                    Name ="Label130"
                                    Caption ="Shelf"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =5775
                            Top =1080
                            Width =245
                            OptionValue =11
                            Name ="Option113"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextAlign =2
                                    Left =5520
                                    Top =780
                                    Width =690
                                    Height =285
                                    Name ="Label114"
                                    Caption ="Size"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =6555
                            Top =1080
                            Width =245
                            OptionValue =7
                            Name ="Option71"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =6360
                                    Top =780
                                    Width =465
                                    Height =285
                                    Name ="Label72"
                                    Caption ="Price"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =7215
                            Top =1080
                            Width =245
                            OptionValue =8
                            Name ="Option66"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =6960
                                    Top =780
                                    Width =765
                                    Height =285
                                    Name ="Label67"
                                    Caption ="Ext Price"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =7875
                            Top =1080
                            Width =245
                            OptionValue =5
                            Name ="Option62"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =7740
                                    Top =780
                                    Width =420
                                    Height =285
                                    Name ="Label63"
                                    Caption ="Cost"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =8480
                            Top =1080
                            Width =245
                            OptionValue =6
                            Name ="Option64"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =8220
                                    Top =780
                                    Width =720
                                    Height =285
                                    Name ="Label65"
                                    Caption ="Ext Cost"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =9075
                            Top =1080
                            Width =245
                            OptionValue =10
                            Name ="Option92"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =9000
                                    Top =780
                                    Width =270
                                    Height =285
                                    Name ="Label93"
                                    Caption ="GP"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =9675
                            Top =1080
                            Width =245
                            OptionValue =9
                            Name ="Option68"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =9360
                                    Top =780
                                    Width =690
                                    Height =285
                                    Name ="Label69"
                                    Caption ="Repacks"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =10320
                            Top =1080
                            Width =245
                            OptionValue =12
                            Name ="Option126"
                            DefaultValue ="10"

                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextAlign =2
                                    Left =10080
                                    Top =780
                                    Width =540
                                    Height =285
                                    Name ="Label127"
                                    Caption ="Labels"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =7860
                    Top =1320
                    TabIndex =19
                    Name ="chkDisplayCost"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6540
                    Top =1320
                    TabIndex =20
                    Name ="chkDisplayPrice"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =255
                    Left =240
                    Top =1260
                    Width =1125
                    Height =240
                    Name ="lblShowColumns"
                    Caption ="Show Columns"
                    FontName ="Trebuchet MS"
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =1620
                    Top =1320
                    TabIndex =22
                    Name ="chkDisplayLine"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =2400
                    Top =1320
                    TabIndex =23
                    Name ="chkDisplayQuantity"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =3180
                    Top =1320
                    TabIndex =24
                    Name ="chkDisplayItemID"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =3960
                    Top =1320
                    TabIndex =25
                    Name ="chkDisplayDescription"
                    StatusBarText ="d"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =7200
                    Top =1320
                    TabIndex =26
                    Name ="chkDisplayExtPrice"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =8465
                    Top =1320
                    TabIndex =27
                    Name ="chkDisplayExtCost"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =9660
                    Top =1320
                    TabIndex =28
                    Name ="chkDisplayRepack"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =9060
                    Top =1320
                    TabIndex =29
                    Name ="chkDisplayGP"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =5760
                    Top =1320
                    TabIndex =30
                    Name ="chkDisplaySize"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =72
                    Left =8700
                    Top =375
                    Width =185
                    TabIndex =31
                    Name ="chkShowChart"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    UnicodeAccessKey =104

                    LayoutCachedLeft =8700
                    LayoutCachedTop =375
                    LayoutCachedWidth =8885
                    LayoutCachedHeight =615
                    Begin
                        Begin Label
                            SpecialEffect =2
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =1
                            Left =8925
                            Top =345
                            Width =795
                            Height =255
                            BackColor =65280
                            ForeColor =-2147483640
                            Name ="Label121"
                            Caption ="C&hart"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =8925
                            LayoutCachedTop =345
                            LayoutCachedWidth =9720
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =12060
                    Top =120
                    Width =1290
                    Height =255
                    TabIndex =33
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =12060
                    LayoutCachedTop =120
                    LayoutCachedWidth =13350
                    LayoutCachedHeight =375
                End
                Begin Subform
                    OverlapFlags =247
                    Left =180
                    Top =3360
                    Width =13440
                    Height =2210
                    Name ="PurchaseOrdersOnFileSub"
                    SourceObject ="Form.PurchaseOrdersOnFileSub"
                    LinkChildFields ="PurchaseOrderID;VendorID"
                    LinkMasterFields ="ID;VendorID"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10305
                    Top =1320
                    TabIndex =34
                    Name ="chkDisplayPrintLabels"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =7560
                    Top =2010
                    TabIndex =35
                    Name ="chkSkipQuantityOnNewItems"
                    DefaultValue ="0"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =7785
                            Top =1980
                            Width =3885
                            Height =435
                            FontWeight =600
                            ForeColor =0
                            Name ="Label136"
                            Caption ="SCAN MODE: On Inserted PO Lines Get Code Number First and order 1 case quantity"
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =255
                    Left =7800
                    Top =2520
                    Width =3165
                    Height =255
                    BackColor =16777215
                    ForeColor =1279872587
                    Name ="Label8"
                    Caption ="OLAS Website (www.olcc.state.or.us/olas)"
                    FontName ="Tahoma"
                    HyperlinkAddress ="http://www.olcc.state.or.us/olas"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =12060
                    Top =495
                    Width =1290
                    Height =255
                    TabIndex =36
                    Name ="txtSetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =12060
                    LayoutCachedTop =495
                    LayoutCachedWidth =13350
                    LayoutCachedHeight =750
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =3120
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =215
                    Left =120
                    Top =900
                    Width =11100
                    Height =2100
                    Name ="PurchaseOrdersOnFileSubItemMovement"
                    SourceObject ="Form.PurchaseOrdersOnFileSubItemMovement"

                End
                Begin Subform
                    OverlapFlags =215
                    Left =120
                    Top =120
                    Width =13560
                    Height =720
                    TabIndex =1
                    Name ="PurchaseOrdersOnFileSubItemSetup"
                    SourceObject ="Form.PurchaseOrdersOnFileSubItemSetup"

                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =60
                    Top =60
                    Width =13620
                    Height =3060
                    Name ="Box75"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11340
                    Top =900
                    Width =1140
                    Height =435
                    TabIndex =2
                    Name ="cmdVendorPOHistory"
                    Caption ="Previous PO's"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11340
                    LayoutCachedTop =900
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =1335
                    Overlaps =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =12720
                    Top =2460
                    Width =666
                    Height =606
                    TabIndex =3
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"

                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =80
                    Left =11340
                    Top =2460
                    Width =576
                    Height =576
                    TabIndex =4
                    Name ="cmdPrintDocs"
                    Caption ="cmdPreview"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    UnicodeAccessKey =80

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =11880
                            Top =2460
                            Width =780
                            Height =405
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label9"
                            Caption ="&Print / Post"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =223
                    AccessKey =69
                    Left =11340
                    Top =1800
                    Width =576
                    Height =576
                    TabIndex =5
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =11940
                            Top =2040
                            Width =630
                            Height =210
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =82
                    Left =12540
                    Top =900
                    Width =1035
                    Height =435
                    TabIndex =6
                    Name ="cmdRefresh"
                    Caption ="&Refresh"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =82

                    LayoutCachedLeft =12540
                    LayoutCachedTop =900
                    LayoutCachedWidth =13575
                    LayoutCachedHeight =1335
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12496
                    Top =1370
                    Width =1086
                    Height =426
                    TabIndex =7
                    Name ="cmdShowDetail"
                    Caption ="Show Detail"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Tag ="20192"
                    ControlTipText ="Open Form"

                    LayoutCachedLeft =12496
                    LayoutCachedTop =1370
                    LayoutCachedWidth =13582
                    LayoutCachedHeight =1796
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12637
                    Top =1889
                    Width =900
                    Height =405
                    TabIndex =8
                    Name ="cmdFixUPC"
                    Caption ="Fix  UPC "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12637
                    LayoutCachedTop =1889
                    LayoutCachedWidth =13537
                    LayoutCachedHeight =2294
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =11340
                    Top =1380
                    Width =1020
                    Height =420
                    TabIndex =9
                    ForeColor =-2147483630
                    Name ="cmdPrintLabels"
                    Caption ="Print Labels"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrdersOnFile.cls"
