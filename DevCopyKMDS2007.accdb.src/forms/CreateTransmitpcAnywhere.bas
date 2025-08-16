Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =36
    Left =2280
    Top =3360
    Right =13230
    Bottom =9900
    RecSrcDt = Begin
        0xddc6ef94c7a2e240
    End
    RecordSource ="Sysfile"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin Section
            Height =8580
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =937
                    Width =6900
                    Height =1875
                    FontSize =36
                    Name ="Label0"
                    Caption ="Create Transmit File pcAnyWhere"
                    FontName ="Trebuchet MS"
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =60
                    Width =8580
                    Height =8520
                    Name ="boxScreen"
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =1440
                    Top =3180
                    Width =6360
                    Height =4140
                    FontSize =10
                    TabIndex =4
                    BackColor =13434828
                    Name ="lstInvoices"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PurchaseOrder.ID, Vendor.Name, Format([OrderDate],\"mm/dd\") AS [Date], P"
                        "urchaseOrder.OrderCount AS InvNumber, qryPurchaseOrderLineTotals.Quantity, qryPu"
                        "rchaseOrderLineTotals.Dollars FROM (PurchaseOrder INNER JOIN Vendor ON PurchaseO"
                        "rder.VendorID=Vendor.ID) INNER JOIN qryPurchaseOrderLineTotals ON PurchaseOrder."
                        "ID=qryPurchaseOrderLineTotals.PurchaseOrderID ORDER BY Vendor.Name, Format([Orde"
                        "rDate],\"mm/dd\"); "
                    ColumnWidths ="0;2160;720;720;1080;1080"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =1440
                            Top =2880
                            Width =4155
                            Height =240
                            Name ="Outstanding Orders_Label"
                            Caption ="To create a PO Transmit File.  First select an order below"
                            EventProcPrefix ="Outstanding_Orders_Label"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =7380
                    Width =5520
                    Height =300
                    FontSize =10
                    TabIndex =5
                    BackColor =16777164
                    Name ="txtSelectedInvoice"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =120
                            Top =7380
                            Width =1260
                            Height =240
                            Name ="Label34"
                            Caption ="Selected Invoice"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =84
                    Left =1560
                    Top =7860
                    Width =2760
                    Height =405
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdEditThisOrder"
                    Caption ="Create &Transmit File"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =84

                    Overlaps =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =67
                    Left =6180
                    Top =7860
                    Width =666
                    Height =606
                    TabIndex =7
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =6900
                            Top =7860
                            Width =585
                            Height =405
                            Name ="Label65"
                            Caption ="&Close Form"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2580
                    Top =1920
                    Width =540
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =1380
                            Top =1920
                            Width =1110
                            Height =240
                            Name ="Label30"
                            Caption ="Agent Number:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3180
                    Top =1920
                    Width =2220
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="Name"

                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =2580
                    Top =2250
                    TabIndex =2
                    Name ="chkSendInventory"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2805
                            Top =2220
                            Width =2595
                            Height =240
                            Name ="Label33"
                            Caption ="Send the inventory"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =2580
                    Top =2550
                    TabIndex =3
                    Name ="chkSendPurchaseOrder"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2805
                            Top =2520
                            Width =2595
                            Height =240
                            Name ="Label35"
                            Caption ="Send a Purchase Order"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CreateTransmitpcAnywhere.cls"
