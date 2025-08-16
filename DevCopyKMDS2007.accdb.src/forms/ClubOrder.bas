Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =10
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13560
    DatasheetFontHeight =10
    ItemSuffix =160
    Left =435
    Top =150
    Right =14370
    Bottom =9015
    Filter ="[ID]=3397"
    RecSrcDt = Begin
        0x7b97ef4f0ef0e440
    End
    RecordSource ="SELECT OrdHead.* FROM OrdHead ORDER BY OrdHead.InvoiceNo;"
    Caption ="Dispenser Orders KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Softwar"
        "e ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS So"
        "ftware ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KM"
        "DS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software "
        "©  KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    MenuBar ="KMDS"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Tahoma"
    OnActivate ="[Event Procedure]"
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
            Height =660
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    TextAlign =2
                    Left =3540
                    Top =240
                    Width =1020
                    Height =330
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    BackColor =8453888
                    ForeColor =0
                    Name ="txtTotalQ"
                    ControlSource ="InvQty"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =3540
                            Width =1020
                            Height =240
                            Name ="Label41"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =4980
                    Top =240
                    Width =1260
                    Height =330
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =8453888
                    ForeColor =0
                    Name ="txtTotalPrice"
                    ControlSource ="InvDollars"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =240
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =570
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =4980
                            Width =1260
                            Height =240
                            Name ="Label45"
                            Caption =" Retail Price"
                            LayoutCachedLeft =4980
                            LayoutCachedWidth =6240
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =6240
                    Top =240
                    Width =1200
                    Height =330
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =8453888
                    ForeColor =0
                    Name ="txtTotalDiscount"
                    ControlSource ="DisDollars"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =240
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =570
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =6240
                            Width =1200
                            Height =240
                            Name ="Label42"
                            Caption ="Discount"
                            LayoutCachedLeft =6240
                            LayoutCachedWidth =7440
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =95
                    TextAlign =2
                    Left =7440
                    Top =240
                    Width =1260
                    Height =330
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =8453888
                    ForeColor =0
                    Name ="txtTotalNet"
                    ControlSource ="OrdDollars"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =240
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =570
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =7440
                            Width =1260
                            Height =240
                            Name ="Label44"
                            Caption ="Net Price"
                            LayoutCachedLeft =7440
                            LayoutCachedWidth =8700
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10980
                    Top =150
                    ColumnOrder =4
                    TabIndex =4
                    Name ="chkShowOrderInfo"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="True"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =11205
                            Top =120
                            Width =1365
                            Height =240
                            Name ="Label94"
                            Caption ="Show Order Info"
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =87
                    ColumnCount =22
                    ListRows =25
                    ListWidth =7920
                    Left =120
                    Top =240
                    Width =3300
                    Height =330
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =5
                    BackColor =16777088
                    Name ="cmbOrderID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OrdHead.ID, Club.Name, OrdHead.ClubID, OrdHead.InvoiceNo, OrdHead.OrdDate"
                        ", qryClubOrderLineTotals.ExtPrice, Club.Contact, Club.Address1, Club.Address2, ["
                        "City] & \", \" & [St] & \"  \" & [Zip] AS CitySt, Club.Phone1, Club.NonLiquor, q"
                        "ryClubOrderLineTotals.ExtCost, qryClubOrderLineTotals.LiquorDiscountRate, qryClu"
                        "bOrderLineTotals.ExtDiscount AS Discount, [ExtPrice]-[ExtDiscount] AS ExtNetPric"
                        "e, Club.ApplyDiscount, Club.MaintainParLevels, qryClubOrderLineTotals.LiquorExtP"
                        "rice, Club.City, Club.St, Club.Zip, Club.FintechAccount FROM Club INNER JOIN (Or"
                        "dHead LEFT JOIN qryClubOrderLineTotals ON OrdHead.ID = qryClubOrderLineTotals.Or"
                        "dheadID) ON Club.ID = OrdHead.ClubID ORDER BY Club.Name;"
                    ColumnWidths ="0;2880;0;1440;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =120
                            Width =390
                            Height =240
                            Name ="Label24"
                            Caption ="Club"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =8685
                    Width =2205
                    Height =570
                    FontSize =20
                    FontWeight =700
                    BackColor =16777215
                    ForeColor =8388608
                    Name ="Label50"
                    Caption ="Dispenser"
                    FontName ="Trebuchet MS"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12720
                    Top =120
                    Width =795
                    Height =405
                    TabIndex =6
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6540
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =4890
                    Top =300
                    Width =2190
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    BackColor =8454143
                    Name ="Order"
                    ControlSource ="Order"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3780
                            Top =300
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
                    OverlapFlags =93
                    TextAlign =2
                    Left =4890
                    Top =900
                    Width =2190
                    Height =255
                    ColumnWidth =900
                    TabIndex =5
                    BackColor =8454143
                    Name ="InvoiceNo"
                    ControlSource ="InvoiceNo"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3780
                            Top =900
                            Width =960
                            Height =240
                            Name ="InvoiceNo_Label"
                            Caption ="Invoice No."
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =4890
                    Top =600
                    Width =2190
                    Height =255
                    ColumnWidth =825
                    TabIndex =3
                    BackColor =8454143
                    Name ="OrdDate"
                    ControlSource ="OrdDate"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3780
                            Top =600
                            Width =915
                            Height =240
                            Name ="OrdDate_Label"
                            Caption ="Order Date"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =4890
                    Top =1200
                    Width =2190
                    Height =255
                    ColumnWidth =825
                    TabIndex =7
                    BackColor =8454143
                    Name ="InvDate"
                    ControlSource ="InvDate"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3780
                            Top =1200
                            Width =1050
                            Height =240
                            Name ="InvDate_Label"
                            Caption ="Invoice Date"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =180
                    Top =2940
                    Width =13320
                    Height =3485
                    TabIndex =11
                    Name ="ClubOrderSub"
                    SourceObject ="Form.ClubOrderSub"
                    LinkChildFields ="OrdheadID"
                    LinkMasterFields ="ID"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =2940
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =6425
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =300
                    Width =3300
                    BackColor =8454143
                    Name ="txtClubName"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =900
                    Width =3300
                    TabIndex =4
                    BackColor =8454143
                    Name ="txtClubAddress1"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =1500
                    Width =1920
                    TabIndex =8
                    BackColor =8454143
                    Name ="txtClubCityStZip"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    Left =2220
                    Top =1500
                    Width =1320
                    TabIndex =9
                    BackColor =8454143
                    Name ="txtClubPhone"
                    FontName ="Trebuchet MS"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =1200
                    Width =3300
                    TabIndex =6
                    BackColor =8454143
                    Name ="txtClubAddress2"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =240
                    Top =600
                    Width =3300
                    TabIndex =2
                    BackColor =8454143
                    Name ="txtClubContact"
                    FontName ="Trebuchet MS"

                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =60
                    Width =3540
                    Height =1860
                    Name ="Box33"
                End
                Begin Label
                    OverlapFlags =255
                    Left =120
                    Top =60
                    Width =570
                    Height =210
                    Name ="Label34"
                    Caption ="Sold to"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3660
                    Top =60
                    Width =3540
                    Height =1860
                    Name ="Box35"
                End
                Begin Label
                    OverlapFlags =255
                    Left =3660
                    Top =60
                    Width =870
                    Height =285
                    Name ="Label36"
                    Caption ="Order Info"
                    FontName ="Trebuchet MS"
                End
                Begin Rectangle
                    OverlapFlags =127
                    Left =7200
                    Top =60
                    Width =2220
                    Height =1860
                    Name ="Box38"
                End
                Begin Label
                    OverlapFlags =255
                    Left =7200
                    Top =60
                    Width =780
                    Height =210
                    Name ="Label39"
                    Caption ="Totals"
                End
                Begin OptionGroup
                    TabStop = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    Left =240
                    Top =1980
                    Width =11880
                    Height =1276
                    TabIndex =10
                    BackColor =16777166
                    Name ="optOrderBy"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =300
                            Top =2220
                            Width =585
                            Height =240
                            Name ="Label51"
                            Caption ="Sort By"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1080
                            Top =2368
                            OptionValue =1
                            Name ="Option53"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =900
                                    Top =2040
                                    Width =540
                                    Height =240
                                    Name ="Label54"
                                    Caption ="LineID"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1920
                            Top =2368
                            OptionValue =4
                            Name ="Option59"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =1680
                                    Top =2040
                                    Width =660
                                    Height =240
                                    Name ="Label60"
                                    Caption ="Quantity"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2820
                            Top =2368
                            OptionValue =2
                            Name ="Option55"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =2580
                                    Top =2040
                                    Width =600
                                    Height =240
                                    Name ="Label56"
                                    Caption ="Code #"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3600
                            Top =2368
                            OptionValue =3
                            Name ="Option57"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =3240
                                    Top =2040
                                    Width =870
                                    Height =240
                                    Name ="Label58"
                                    Caption ="Description"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =4320
                            Top =2368
                            OptionValue =9
                            Name ="Option87"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =4200
                                    Top =2040
                                    Width =825
                                    Height =240
                                    Name ="Label88"
                                    Caption ="Size; code"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =5220
                            Top =2368
                            OptionValue =10
                            Name ="Option119"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =5100
                                    Top =2040
                                    Width =840
                                    Height =240
                                    Name ="Label120"
                                    Caption ="Size; Desc"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6180
                            Top =2368
                            OptionValue =11
                            Name ="Option134"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6000
                                    Top =2040
                                    Width =1200
                                    Height =240
                                    Name ="Label135"
                                    Caption ="Category;Code"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =7440
                            Top =2368
                            OptionValue =12
                            Name ="Option141"
                            OnGotFocus ="[Event Procedure]"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =7260
                                    Top =2040
                                    Width =780
                                    Height =240
                                    Name ="Label142"
                                    Caption ="Cat;Descr"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8340
                            Top =2368
                            OptionValue =13
                            Name ="Option143"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =8160
                                    Top =2040
                                    Width =780
                                    Height =240
                                    Name ="Label144"
                                    Caption ="Cat;Price"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =9300
                            Top =2368
                            OptionValue =5
                            Name ="Option61"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =9120
                                    Top =2040
                                    Width =435
                                    Height =240
                                    Name ="Label62"
                                    Caption ="Price"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =9780
                            Top =2368
                            OptionValue =6
                            Name ="Option63"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =9600
                                    Top =2040
                                    Width =390
                                    Height =240
                                    Name ="Label64"
                                    Caption ="Cost"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =10260
                            Top =2368
                            OptionValue =7
                            Name ="Option65"
                            OnGotFocus ="[Event Procedure]"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =10020
                                    Top =2040
                                    Width =660
                                    Height =240
                                    Name ="Label66"
                                    Caption ="ExtPrice"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =10920
                            Top =2368
                            OptionValue =8
                            Name ="Option67"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =10740
                                    Top =2040
                                    Width =615
                                    Height =240
                                    Name ="Label68"
                                    Caption ="ExtCost"
                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =60
                    Width =13440
                    Height =6480
                    Name ="BoxBody"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =420
                    Top =2580
                    Width =480
                    Height =240
                    Name ="lblShowColumns"
                    Caption ="Show"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =8940
                    Top =240
                    Width =285
                    Height =255
                    ColumnWidth =225
                    TabIndex =12
                    Name ="Status"
                    ControlSource ="Status"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            Left =8340
                            Top =240
                            Width =585
                            Height =240
                            Name ="Status_Label"
                            Caption ="Status"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =8040
                    Top =300
                    Width =320
                    TabIndex =13
                    Name ="PaidFlag"
                    ControlSource ="PaidFlag"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =7260
                            Top =300
                            Width =750
                            Height =240
                            Name ="PaidFlag_Label"
                            Caption ="PaidFlag"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    Left =7920
                    Top =600
                    Width =720
                    Height =255
                    ColumnWidth =900
                    TabIndex =14
                    Name ="ClubID"
                    ControlSource ="ClubID"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            TextAlign =3
                            Left =7260
                            Top =600
                            Width =555
                            Height =240
                            Name ="ClubID_Label"
                            Caption ="ClubID"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    Left =8700
                    Top =600
                    Width =600
                    Height =255
                    ColumnWidth =1215
                    TabIndex =15
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =2880
                    Width =13440
                    Height =3660
                    Name ="boxSubFormNormal"
                    LayoutCachedLeft =120
                    LayoutCachedTop =2880
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =6540
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =8760
                    Top =1020
                    TabIndex =16
                    Name ="chkNonLiquor"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =7425
                            Top =960
                            Width =1305
                            Height =240
                            Name ="Label109"
                            Caption ="Non-Liquor Sales"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10380
                    Top =420
                    Width =945
                    TabIndex =17
                    BackColor =8454143
                    Name ="txtLiquorDollars"
                    ControlSource ="LiquorDollars"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =10380
                    LayoutCachedTop =420
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9735
                            Top =420
                            Width =600
                            Height =285
                            Name ="Label111"
                            Caption ="Retail"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10380
                    Top =720
                    Width =930
                    TabIndex =18
                    BackColor =8454143
                    Name ="txtLiquorDiscount"
                    ControlSource ="LiquorDiscount"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =10380
                    LayoutCachedTop =720
                    LayoutCachedWidth =11310
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9540
                            Top =720
                            Width =735
                            Height =285
                            Name ="Label112"
                            Caption ="Discount"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =9540
                            LayoutCachedTop =720
                            LayoutCachedWidth =10275
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =9480
                    Top =1560
                    Width =705
                    TabIndex =19
                    Name ="cmdResort"
                    Caption ="Resort"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10260
                    Top =1560
                    Width =1680
                    TabIndex =20
                    Name ="cmdSaveOrderGuide"
                    Caption ="Save Order Guide"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =6480
                    Top =2340
                    Width =2715
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =16777215
                    ForeColor =255
                    Name ="lblRetr"
                    Caption ="Inserting Line"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    Left =4920
                    Top =1500
                    Width =660
                    TabIndex =21
                    BackColor =8454143
                    Name ="txtDiscountRate"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    Left =5700
                    Top =1500
                    Width =1320
                    TabIndex =22
                    BackColor =8454143
                    Name ="txtNetPrice"
                    FontName ="Trebuchet MS"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =8760
                    Top =1260
                    TabIndex =23
                    Name ="chkParLevels"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =7260
                            Top =1200
                            Width =1470
                            Height =240
                            Name ="Label128"
                            Caption ="Maintain Par Levels"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =8760
                    Top =1500
                    TabIndex =24
                    Name ="chkApplyDiscounts"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =7590
                            Top =1440
                            Width =1140
                            Height =240
                            Name ="Label130"
                            Caption ="Apply Discount"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1080
                    Top =2580
                    TabIndex =25
                    Name ="chkLineID"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =2820
                    Top =2580
                    TabIndex =26
                    Name ="chkItemID"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =3600
                    Top =2580
                    TabIndex =27
                    Name ="chkDesc"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =10260
                    Top =2580
                    TabIndex =28
                    Name ="chkExtPrice"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =10920
                    Top =2580
                    TabIndex =29
                    Name ="chkExtCost"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =9300
                    Top =2580
                    TabIndex =30
                    Name ="chkPrice"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =4320
                    Top =2580
                    TabIndex =31
                    Name ="chkSize"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1920
                    Top =2580
                    TabIndex =32
                    Name ="chkQuantity"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =9780
                    Top =2580
                    TabIndex =33
                    Name ="chkCost"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6180
                    Top =2580
                    TabIndex =34
                    Name ="chkCategory"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11445
                    Top =420
                    Width =810
                    TabIndex =35
                    BackColor =8454143
                    Name ="txtNonLiquorDollars"
                    ControlSource ="NonLiquorDollars"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =10455
                            Top =120
                            Width =2685
                            Height =255
                            Name ="Label154"
                            Caption ="Liquor    Non Liquor          Total "
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11445
                    Top =720
                    Width =810
                    TabIndex =36
                    BackColor =8454143
                    Name ="txtNonLiquorDiscount"
                    ControlSource ="NonLiquorDiscount"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11415
                    Top =1080
                    Width =810
                    TabIndex =37
                    BackColor =8453888
                    Name ="txtInvQty"
                    ControlSource ="InvQty"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9600
                            Top =1125
                            Width =735
                            Height =285
                            Name ="Label156"
                            Caption ="Total"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12315
                    Top =420
                    Width =975
                    TabIndex =38
                    BackColor =12909560
                    Name ="txtInvDollars"
                    ControlSource ="InvDollars"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =12315
                    LayoutCachedTop =420
                    LayoutCachedWidth =13290
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12315
                    Top =720
                    Width =990
                    TabIndex =39
                    BackColor =12909560
                    Name ="txtDisDollars"
                    ControlSource ="DisDollars"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =12315
                    LayoutCachedTop =720
                    LayoutCachedWidth =13305
                    LayoutCachedHeight =960
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12315
                    Top =1080
                    Width =990
                    FontWeight =700
                    TabIndex =40
                    BackColor =8453888
                    Name ="txtOrdDollars"
                    ControlSource ="OrdDollars"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =12315
                    LayoutCachedTop =1080
                    LayoutCachedWidth =13305
                    LayoutCachedHeight =1320
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =11940
                    Top =1620
                    TabIndex =41
                    Name ="chkSkipQuantityOnNewItems"
                    DefaultValue ="0"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =12180
                            Top =1560
                            Width =1320
                            Height =1260
                            FontWeight =600
                            ForeColor =0
                            Name ="Label136"
                            Caption ="SCAN MODE: On Inserted Lines Get Code Number First and order 1."
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1626
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =1
                    ListWidth =11520
                    Left =1440
                    Top =60
                    Width =9180
                    Height =480
                    BackColor =8454143
                    Name ="cmbClubNotes"
                    ControlSource ="ClubID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Club.ID, Club.ClubNotes FROM Club; "
                    ColumnWidths ="0;11520"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =60
                            Width =600
                            Height =480
                            Name ="Label98"
                            Caption ="Club Notes:"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1440
                    Top =600
                    Width =9120
                    Height =603
                    TabIndex =1
                    Name ="OrderNotes"
                    ControlSource ="OrderNotes"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =600
                            Width =1320
                            Height =600
                            Name ="Label101"
                            Caption ="Order Notes (prints on customer Invoice)"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10740
                    Top =60
                    Height =480
                    TabIndex =2
                    Name ="cmdPurchasesByItem"
                    Caption ="Purchases by Item"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =60
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =540
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =12240
                    Top =600
                    Width =1200
                    TabIndex =3
                    Name ="cmdDelOrder"
                    Caption ="&Delete Order"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =12240
                    LayoutCachedTop =600
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =960
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =12180
                    Top =1020
                    Width =666
                    Height =546
                    TabIndex =4
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
                    UnicodeAccessKey =67

                    LayoutCachedLeft =12180
                    LayoutCachedTop =1020
                    LayoutCachedWidth =12846
                    LayoutCachedHeight =1566
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =12900
                            Top =1020
                            Width =585
                            Height =405
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label65"
                            Caption ="&Close Form"
                            FontName ="Tahoma"
                            LayoutCachedLeft =12900
                            LayoutCachedTop =1020
                            LayoutCachedWidth =13485
                            LayoutCachedHeight =1425
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =10620
                    Top =900
                    Width =666
                    Height =606
                    TabIndex =5
                    Name ="cmdPrint"
                    Caption ="cmdPreview"
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
                    UnicodeAccessKey =83

                    LayoutCachedLeft =10620
                    LayoutCachedTop =900
                    LayoutCachedWidth =11286
                    LayoutCachedHeight =1506
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =11280
                            Top =960
                            Width =780
                            Height =480
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label9"
                            Caption ="&Save and Print"
                            FontName ="Tahoma"
                            LayoutCachedLeft =11280
                            LayoutCachedTop =960
                            LayoutCachedWidth =12060
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    Left =1440
                    Top =1200
                    Width =9120
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblMapQuest"
                    Caption ="Click here for &Google Map to Bugatti's"
                    ControlTipText ="http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=,+Hillsboro,"
                        "+Or+97006&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
                    HyperlinkAddress ="http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=,+Hillsboro,"
                        "+Or+97006&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12240
                    Top =60
                    Width =1260
                    Height =480
                    TabIndex =6
                    Name ="cmdPreviousInvoiceReport"
                    Caption ="Previous Invoices"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12240
                    LayoutCachedTop =60
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =540
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubOrder.cls"
