Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15600
    RowHeight =270
    DatasheetFontHeight =10
    ItemSuffix =62
    Left =3345
    Top =5415
    Right =15795
    Bottom =7245
    DatasheetBackColor =13434828
    ShortcutMenuBar ="KMDSPopup"
    RecSrcDt = Begin
        0x9448c91db06be340
    End
    RecordSource ="SELECT AgentTransferLines.ID, AgentTransferLines.AgentTransferHeadID, AgentTrans"
        "ferLines.LineID, AgentTransferLines.ItemID, AgentTransferLines.QuantityShipped, "
        "AgentTransferLines.Price, AgentTransferLines.Cost, AgentTransferLines.SetupDate,"
        " AgentTransferLines.ChangeDate, [QuantityShipped]*[Price] AS ExtPrice, [Quantity"
        "Shipped]*[Cost] AS ExtCost, IIf([Price]>0,([Price]-[Cost])/[Price],0) AS GP, Age"
        "ntTransferLines.QuantityOrdered, AgentTransferLines.ClaimCode FROM AgentTransfer"
        "Lines; "
    Caption ="AgentsTranfersSub"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =255
    DatasheetBackColor12 =13434828
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
            KeepTogether = NotDefault
            Height =285
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =180
                    Width =540
                    Height =240
                    Name ="ID_Label"
                    Caption ="ID"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =5820
                    Width =600
                    Height =240
                    Name ="QuantityOrdered_Label"
                    Caption ="QuantityOrdered"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =3
                    Left =6420
                    Width =1140
                    Height =240
                    Name ="Price_Label"
                    Caption ="Price"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =4800
                    Width =720
                    Height =255
                    ColumnOrder =0
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3960
                    Width =660
                    Height =255
                    ColumnOrder =1
                    TabIndex =1
                    Name ="AgentTransferHeadID"
                    ControlSource ="AgentTransferHeadID"
                    FontName ="Trebuchet MS"

                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =720
                    Width =420
                    Height =285
                    Name ="Label31"
                    Caption ="Item"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    Left =7560
                    Width =1200
                    Height =240
                    Name ="Label32"
                    Caption ="Price"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =2880
                    Width =480
                    Height =255
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =13434828
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityOrdered"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
            End
        End
        Begin Section
            Height =1140
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Width =360
                    Height =255
                    ColumnWidth =465
                    ColumnOrder =0
                    BackColor =13434828
                    Name ="LineID"
                    ControlSource ="LineID"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    Left =420
                    Width =480
                    Height =255
                    ColumnWidth =690
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    Name ="QuantityShipped"
                    ControlSource ="QuantityShipped"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Blank or Null is not allowed here; try using the number '0'."
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =300
                            Width =660
                            Height =210
                            Name ="Label42"
                            Caption ="Qty"
                        End
                    End
                End
                Begin TextBox
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =7140
                    Width =1140
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =6
                    TabIndex =4
                    BackColor =13434828
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8340
                    Width =1140
                    Height =255
                    ColumnWidth =1000
                    ColumnOrder =7
                    TabIndex =5
                    BackColor =16776960
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =9
                    ListRows =25
                    ListWidth =10080
                    Left =960
                    Width =1080
                    Height =255
                    ColumnWidth =855
                    ColumnOrder =2
                    TabIndex =2
                    BoundColumn =8
                    BackColor =13434828
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vwUPCItemsWithCurrentPricing.LenID, vwUPCItemsWithCurrentPricing.UPCCode,"
                        " vwUPCItemsWithCurrentPricing.Description, vwUPCItemsWithCurrentPricing.OnShelf,"
                        " vwUPCItemsWithCurrentPricing.Cost, vwUPCItemsWithCurrentPricing.Price, vwUPCIte"
                        "msWithCurrentPricing.ItemVendorID AS VID, vwUPCItemsWithCurrentPricing.UnitsPerC"
                        "ase, vwUPCItemsWithCurrentPricing.ID AS ItemID FROM vwUPCItemsWithCurrentPricing"
                        " ORDER BY vwUPCItemsWithCurrentPricing.LenID, vwUPCItemsWithCurrentPricing.UPCCo"
                        "de; "
                    ColumnWidths ="0;1440;3600;1440;1440;1440;1440;1440;1440"
                    StatusBarText ="Double click to EDIT this Item."
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial Narrow"
                    OnNotInList ="[Event Procedure]"
                    ShortcutMenuBar ="KMDSPopup"
                    ControlTipText ="Double click to EDIT this Item."

                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =9
                    ListRows =25
                    ListWidth =10080
                    Left =2100
                    Width =4140
                    Height =255
                    ColumnWidth =3465
                    ColumnOrder =4
                    TabIndex =3
                    BoundColumn =8
                    BackColor =13434828
                    Name ="ItemDescription"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vwUPCItemsWithCurrentPricing.Description, vwUPCItemsWithCurrentPricing.Le"
                        "nID, vwUPCItemsWithCurrentPricing.UPCCode, vwUPCItemsWithCurrentPricing.OnShelf,"
                        " vwUPCItemsWithCurrentPricing.Cost, vwUPCItemsWithCurrentPricing.Price, vwUPCIte"
                        "msWithCurrentPricing.ItemVendorID AS VID, vwUPCItemsWithCurrentPricing.UnitsPerC"
                        "ase, vwUPCItemsWithCurrentPricing.ID AS ItemID FROM vwUPCItemsWithCurrentPricing"
                        " ORDER BY vwUPCItemsWithCurrentPricing.Description, vwUPCItemsWithCurrentPricing"
                        ".LenID, vwUPCItemsWithCurrentPricing.UPCCode; "
                    ColumnWidths ="3600;0;1440;1440;1440;1440;1440;1440"
                    StatusBarText ="Double click to EDIT this Item."
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"
                    ControlTipText ="Double click to EDIT this Item."

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =25
                    ListWidth =2880
                    Left =6300
                    Width =720
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =5
                    TabIndex =6
                    BoundColumn =1
                    BackColor =13434828
                    Name ="Size"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vwUPCItemsWithCurrentPricing.PkgSize, vwUPCItemsWithCurrentPricing.ID AS "
                        "ItemID FROM vwUPCItemsWithCurrentPricing; "
                    ColumnWidths ="1440;1440"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =25
                    ListWidth =2880
                    Left =4920
                    Top =360
                    Width =720
                    Height =255
                    ColumnWidth =735
                    ColumnOrder =11
                    TabIndex =7
                    BoundColumn =1
                    BackColor =13434828
                    Name ="On Shelf"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vwUPCItemsWithCurrentPricing.OnShelf, vwUPCItemsWithCurrentPricing.ID AS "
                        "ItemID FROM vwUPCItemsWithCurrentPricing; "
                    ColumnWidths ="1440;1440"
                    FontName ="Trebuchet MS"
                    EventProcPrefix ="On_Shelf"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =9480
                    Top =420
                    ColumnWidth =1500
                    ColumnOrder =3
                    TabIndex =8
                    Name ="ClaimCode"
                    ControlSource ="ClaimCode"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Left([ReportName],15) AS ClaimCode FROM ReportForms WHERE (((ReportForms."
                        "ReportType)=\"Claims\")) ORDER BY Left([ReportName],15); "
                    ColumnWidths ="1440;1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8040
                            Top =420
                            Width =915
                            Height =245
                            Name ="Claim Code_Label"
                            Caption ="Claim Code"
                            EventProcPrefix ="Claim_Code_Label"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =645
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    Left =5760
                    Top =60
                    Width =540
                    Height =255
                    Name ="txtSumQuantity"
                    ControlSource ="=Sum([QuantityShipped])"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =7320
                    Top =60
                    Width =1380
                    TabIndex =1
                    Name ="txtSumPrice"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =9720
                    Top =60
                    Width =1380
                    TabIndex =2
                    Name ="txtSumCost"
                    ControlSource ="=Sum([ExtCost])"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =420
                    Top =240
                    Width =795
                    Height =405
                    TabIndex =3
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                End
            End
        End
    End
End
CodeBehindForm
' See "AgentsTransfersSub.cls"
