Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10920
    DatasheetFontHeight =10
    ItemSuffix =78
    Filter ="((Item.ID)=1302)"
    RecSrcDt = Begin
        0x1bd0e5781b8be240
    End
    RecordSource ="Item"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =705
                    Top =5
                    Width =735
                    Height =300
                    FontSize =10
                    BackColor =8453888
                    ForeColor =0
                    Name ="OnReserve"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OrderLine.ItemID, Sum(OrderLine.QuantityShipped) AS SumOfQuantityShipped "
                        "FROM OrderLine GROUP BY OrderLine.ItemID; "
                    ColumnWidths ="0;1440;1440"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click on this number to see what Order(s) have this item."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Width =675
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label110"
                            Caption ="Reserve"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1020
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =60
                    Top =360
                    Width =2220
                    Height =255
                    ColumnWidth =900
                    BackColor =16777215
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2340
                    Top =360
                    Width =420
                    Height =300
                    ColumnWidth =900
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =16777215
                    ForeColor =16711680
                    Name ="txtAvailOnTheShelf"
                    BeforeUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =2820
                    Top =360
                    Width =360
                    Height =255
                    ColumnWidth =900
                    TabIndex =2
                    BackColor =16777215
                    Name ="OrderAt"
                    ControlSource ="OrderAt"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =3240
                    Top =360
                    Width =360
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    BackColor =16777215
                    Name ="OrderTo"
                    ControlSource ="OrderTo"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =3720
                    Top =360
                    Width =360
                    Height =255
                    ColumnWidth =900
                    TabIndex =4
                    BackColor =16777215
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =4380
                    Top =360
                    Width =360
                    Height =255
                    ColumnWidth =900
                    TabIndex =5
                    BackColor =16777215
                    Name ="UnitsPerOrder"
                    ControlSource ="UnitsPerOrder"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =6720
                    Top =360
                    Width =1020
                    Height =255
                    ColumnWidth =1560
                    TabIndex =9
                    Name ="StockNumber"
                    ControlSource ="StockNumber"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =7860
                    Top =360
                    Width =960
                    Height =255
                    ColumnWidth =825
                    TabIndex =10
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"
                    Format ="Short Date"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Width =1065
                    Height =285
                    Name ="Label58"
                    Caption ="Current Item"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2145
                    Width =675
                    Height =285
                    Name ="Label59"
                    Caption ="On Shelf"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Width =330
                    Height =285
                    Name ="Label60"
                    Caption ="Min"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3300
                    Width =375
                    Height =285
                    Name ="Label61"
                    Caption ="Max"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3720
                    Width =540
                    Height =285
                    Name ="Label62"
                    Caption ="UCase"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4237
                    Width =630
                    Height =285
                    Name ="Label63"
                    Caption ="UOrder"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6698
                    Width =1065
                    Height =285
                    Name ="Label64"
                    Caption ="StockNumber"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7785
                    Width =1050
                    Height =285
                    Name ="Label65"
                    Caption ="DateLastSold"
                    FontName ="Trebuchet MS"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2070
                    Left =8880
                    Top =360
                    Height =255
                    TabIndex =11
                    Name ="cmbVendorID"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Vendor].[ID], [Vendor].[Name] FROM Vendor ORDER BY [Name]; "
                    ColumnWidths ="0;2070"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9000
                            Width =585
                            Height =285
                            Name ="Vendor_Label"
                            Caption ="Vendor"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =360
                    Top =720
                    TabIndex =12
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =720
                            Width =615
                            Height =240
                            Name ="Label72"
                            Caption ="Text71:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4860
                    Top =360
                    Width =720
                    TabIndex =6
                    BackColor =16777215
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4980
                            Width =480
                            Height =285
                            Name ="Label73"
                            Caption ="Price:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =360
                    Width =660
                    TabIndex =8
                    BackColor =16777215
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =6113
                            Width =435
                            Height =285
                            Name ="Label74"
                            Caption ="Cost:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5700
                    Top =360
                    TabIndex =7
                    Name ="TPR"
                    ControlSource ="TPR"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Width =405
                            Height =285
                            Name ="Label76"
                            Caption ="TPR"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10380
                    Top =360
                    Width =540
                    TabIndex =13
                    Name ="AverageDailySales"
                    ControlSource ="AverageDailySales"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10320
                            Width =585
                            Height =285
                            Name ="Label77"
                            Caption ="AvgDS"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "PurchaseOrdersOnFileSubItemSetupxx.cls"
