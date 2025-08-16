Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12968
    DatasheetFontHeight =10
    ItemSuffix =174
    Left =1800
    Top =2490
    Right =15195
    Bottom =10710
    FrozenColumns =4
    HelpContextId =4100
    OnUnload ="[Event Procedure]"
    Tag ="Inventory Value By Code"
    Filter ="(Item.TPR<>0 Or Item.TPR Is Null)"
    ShortcutMenuBar ="KMDSPopup"
    OrderBy ="Item.Description, Item.ID"
    RecSrcDt = Begin
        0xf9248a63689be540
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Item"
    Caption ="Items by Description by Code"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
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
        Begin Subform
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
        Begin Tab
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
        Begin FormHeader
            Height =0
            BackColor =16764057
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4133
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3555
                    Left =1260
                    Top =1740
                    Width =3840
                    Height =300
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="SubCategoryID"
                    ControlSource ="SubCategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SubCategory.ID, SubCategory.Description FROM SubCategory ORDER BY [Descri"
                        "ption]; "
                    ColumnWidths ="0;3555"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =420
                            Top =1740
                            Width =780
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Category_Label"
                            Caption ="Category"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =11520
                    Top =180
                    Width =1320
                    Height =300
                    TabIndex =15
                    ForeColor =255
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =11520
                    LayoutCachedTop =180
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =10800
                            Top =180
                            Width =660
                            Height =300
                            Name ="Label1"
                            Caption ="Item ID"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =10800
                            LayoutCachedTop =180
                            LayoutCachedWidth =11460
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3960
                    Left =6120
                    Top =1020
                    Width =1650
                    Height =300
                    FontSize =10
                    TabIndex =10
                    BoundColumn =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="SizeID"
                    ControlSource ="SizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.Description, Size.ID FROM [Size] ORDER BY Size.Description;"
                    ColumnWidths ="2880;720"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5430
                            Top =1020
                            Width =630
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label3"
                            Caption ="SizeID"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    ListWidth =5400
                    Left =1260
                    Top =1380
                    Width =3840
                    Height =300
                    ColumnWidth =2850
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =8
                    BoundColumn =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [ID] & \" - \" & [SalesDepartment].[Description] AS SD, SalesDepartment.I"
                        "D FROM SalesDepartment ORDER BY SalesDepartment.ID; "
                    ColumnWidths ="4320;720"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =300
                            Top =1380
                            Width =900
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label5"
                            Caption ="Sales Dept."
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =8670
                    Top =1020
                    Width =1920
                    Height =300
                    ColumnWidth =1950
                    FontSize =10
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name FROM Vendor ORDER BY Vendor.Name; "
                    ColumnWidths ="0;3600"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =7995
                            Top =1020
                            Width =615
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label7"
                            Caption ="Vendor"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    ListWidth =4320
                    Left =6150
                    Top =1380
                    Width =1635
                    Height =300
                    ColumnWidth =2235
                    FontSize =10
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="PriceGroupID"
                    ControlSource ="PriceGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.ID, PriceGroup.Description FROM PriceGroup ORDER BY PriceGroup"
                        ".Description; "
                    ColumnWidths ="0;2880"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5130
                            Top =1380
                            Width =960
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label9"
                            Caption ="Price Group"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =1230
                    Top =1020
                    Width =3840
                    Height =300
                    ColumnWidth =2895
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =7
                    ForeColor =-2147483640
                    Name ="Description"
                    ControlSource ="Description"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =180
                            Top =1020
                            Width =960
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label11"
                            Caption ="Description"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =2220
                    Width =810
                    Height =300
                    FontSize =10
                    TabIndex =16
                    ForeColor =-2147483640
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Blank or Null is not allowed here; try using the number '0'."
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =330
                            Top =2220
                            Width =810
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label13"
                            Caption ="Onhand"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =2580
                    Width =810
                    Height =300
                    FontSize =10
                    TabIndex =17
                    ForeColor =-2147483640
                    Name ="OrderAt"
                    ControlSource ="OrderAt"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =375
                            Top =2580
                            Width =765
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label15"
                            Caption ="OrderAt"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =2940
                    Width =810
                    Height =300
                    FontSize =10
                    TabIndex =18
                    ForeColor =-2147483640
                    Name ="OrderTo"
                    ControlSource ="OrderTo"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =360
                            Top =2940
                            Width =780
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label17"
                            Caption ="OrderTo"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =4620
                    Top =2220
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =19
                    ForeColor =-2147483640
                    Name ="Price"
                    ControlSource ="Price"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Retail Selling price"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3720
                            Top =2220
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label21"
                            Caption ="Price"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =4620
                    Top =2580
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =21
                    ForeColor =-2147483640
                    Name ="Cost"
                    ControlSource ="Cost"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Price paid to acquire goods"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3720
                            Top =2580
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label23"
                            Caption ="Cost"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =4620
                    Top =2940
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Deposit"
                    ControlSource ="Deposit"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Bottle and can deposit"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3720
                            Top =2940
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label25"
                            Caption ="Deposit"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =7515
                    Top =2220
                    Width =795
                    Height =300
                    FontSize =10
                    TabIndex =25
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Used to determine repacks"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6360
                            Top =2220
                            Width =1095
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label27"
                            Caption ="UnitsPerCase"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =7515
                    Top =2580
                    Width =795
                    Height =300
                    FontSize =10
                    TabIndex =26
                    ForeColor =-2147483640
                    Name ="UnitsPerOrder"
                    ControlSource ="UnitsPerOrder"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Purchase Order Multiple - Order Quantity will be divisible by this number"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6270
                            Top =2580
                            Width =1185
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label29"
                            Caption ="UnitsPerOrder"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =7515
                    Top =2940
                    Width =795
                    Height =300
                    FontSize =10
                    TabIndex =27
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="UnitsPerPack"
                    ControlSource ="UnitsPerPack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Packs Per Carton (used in Cash Register Transfer Cartons)"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6360
                            Top =2940
                            Width =1095
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label31"
                            Caption ="UnitsPerPack"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =6960
                    Top =3300
                    Width =1350
                    Height =300
                    FontSize =10
                    TabIndex =28
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="SubPack"
                    ControlSource ="SubPack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="UPC Code for Pack  (used in Cash Register Transfer Cartons)"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6255
                            Top =3300
                            Width =645
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label33"
                            Caption ="Packs "
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =4605
                    Top =3300
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="QtyPack"
                    ControlSource ="QtyPack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Used for 6-pack pricing   6 @ $2.99"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3705
                            Top =3300
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label45"
                            Caption ="QtyPack"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =4605
                    Top =3660
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =24
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="QtyPackPrice"
                    ControlSource ="QtyPackPrice"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Used for 6-pack pricing   6 @ $2.99"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3705
                            Top =3660
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label47"
                            Caption =" PackPrice"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =8640
                    Top =1380
                    Width =1950
                    Height =274
                    FontSize =10
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="StockNumber"
                    ControlSource ="StockNumber"
                    StatusBarText ="Set Stock Number to X to prevent it from ordering..."
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ControlTipText ="Set Stock Number to X to prevent it from ordering..."
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =3
                            Left =7785
                            Top =1380
                            Width =795
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label49"
                            Caption ="Stock#"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =960
                    Width =10800
                    Height =1140
                    Name ="Box70"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =2100
                    Width =2100
                    Height =1860
                    Name ="Box71"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3660
                    Top =2100
                    Width =2460
                    Height =1860
                    Name ="Box72"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =6300
                    Top =2100
                    Width =2220
                    Height =1860
                    Name ="Box73"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =8400
                    Top =2100
                    Width =3057
                    Height =1860
                    Name ="Box74"
                    LayoutCachedLeft =8400
                    LayoutCachedTop =2100
                    LayoutCachedWidth =11457
                    LayoutCachedHeight =3960
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =255
                    TextAlign =2
                    Left =5460
                    Top =2580
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =20
                    BackColor =255
                    Name ="txtGP"
                    Format ="Percent"
                    FontName ="Trebuchet MS"

                End
                Begin Subform
                    OverlapFlags =247
                    Left =8580
                    Top =2160
                    Width =2520
                    Height =1740
                    TabIndex =29
                    Name ="UPCItem"
                    SourceObject ="Form.ItemUPC"
                    LinkChildFields ="ItemID"
                    LinkMasterFields ="ID"
                    OnExit ="[Event Procedure]"

                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =67
                    ColumnCount =7
                    ListRows =25
                    ListWidth =9360
                    Left =600
                    Top =180
                    Width =1500
                    Height =300
                    ColumnWidth =1095
                    ColumnOrder =3
                    FontSize =10
                    BackColor =16777088
                    Name ="cmbItemByCode"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByCode"
                    ColumnWidths ="1440;2880;720;720;720;720;360"
                    StatusBarText ="Double click to Start at UPC Code"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click to Start at UPC Code"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =600
                    LayoutCachedTop =180
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =60
                            Top =180
                            Width =480
                            Height =285
                            ForeColor =16711680
                            Name ="cmbItemByCode_Label"
                            Caption ="&Code"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =60
                            LayoutCachedTop =180
                            LayoutCachedWidth =540
                            LayoutCachedHeight =465
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    AccessKey =68
                    ColumnCount =7
                    ListRows =25
                    ListWidth =7560
                    Left =3165
                    Top =180
                    Width =2160
                    Height =285
                    ColumnWidth =1785
                    ColumnOrder =4
                    TabIndex =1
                    BoundColumn =2
                    BackColor =16777088
                    Name ="cmbItemByDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByDescription"
                    ColumnWidths ="2880;720;1440;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =3165
                    LayoutCachedTop =180
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =465
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2160
                            Top =180
                            Width =945
                            Height =285
                            ForeColor =16711680
                            Name ="Description_Label"
                            Caption ="&Description"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =2160
                            LayoutCachedTop =180
                            LayoutCachedWidth =3105
                            LayoutCachedHeight =465
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =85
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9000
                    Left =5820
                    Top =180
                    Width =1410
                    Height =285
                    ColumnOrder =5
                    TabIndex =2
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemByUPC"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByUPC"
                    ColumnWidths ="1440;2880;720;1440;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =85

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5415
                            Top =180
                            Width =375
                            Height =285
                            ForeColor =16711680
                            Name ="ID_Label"
                            Caption ="&UPC"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =120
                    Width =10740
                    Height =780
                    Name ="Box78"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =83
                    ColumnCount =8
                    ListRows =25
                    ListWidth =10080
                    Left =8160
                    Top =180
                    Width =2580
                    Height =300
                    ColumnOrder =6
                    TabIndex =3
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemBySalesDept"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemBySalesDepartmentID"
                    ColumnWidths ="1800;2880;720;1440;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =83

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =7260
                            Top =180
                            Width =840
                            Height =285
                            ForeColor =16711680
                            Name ="Label81"
                            Caption ="&Sales Dept"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =80
                    ColumnCount =8
                    ListRows =25
                    ListWidth =11520
                    Left =3180
                    Top =540
                    Width =2145
                    Height =285
                    ColumnOrder =7
                    TabIndex =4
                    BoundColumn =3
                    BackColor =16777164
                    Name ="cmbItemsByPriceGroup"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2880;2880;720;1440;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click to activate."
                    UnicodeAccessKey =80

                    LayoutCachedLeft =3180
                    LayoutCachedTop =540
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =825
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =2160
                            Top =540
                            Width =945
                            Height =285
                            ForeColor =16711680
                            Name ="Label83"
                            Caption =" &Price Group"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =2160
                            LayoutCachedTop =540
                            LayoutCachedWidth =3105
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =76
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9360
                    Left =5820
                    Top =540
                    Width =1410
                    Height =285
                    ColumnOrder =8
                    TabIndex =5
                    BoundColumn =3
                    BackColor =16777164
                    Name ="cmbItemByLocation"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="720;2880;720;1440;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click to activate."
                    UnicodeAccessKey =76

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =1
                            Left =5400
                            Top =540
                            Width =375
                            Height =285
                            ForeColor =16711680
                            Name ="Label87"
                            Caption ="&Location"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =540
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =89
                    ColumnCount =8
                    ListRows =25
                    ListWidth =10080
                    Left =8160
                    Top =540
                    Width =2580
                    Height =300
                    ColumnOrder =9
                    TabIndex =6
                    BoundColumn =3
                    BackColor =16777164
                    Name ="cmbItemBySubCategoryID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2160;2880;720;1440;720;720;720;360"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double click to activate."
                    UnicodeAccessKey =121

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =7320
                            Top =540
                            Width =780
                            Height =285
                            ForeColor =16711680
                            Name ="Label101"
                            Caption ="Categor&y"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin OptionButton
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10200
                    Top =1740
                    TabIndex =30
                    Name ="optTabSkip"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =8100
                            Top =1740
                            Width =2025
                            Height =240
                            Name ="Label105"
                            Caption ="Stop on Specified Fields"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =1350
                    Top =3300
                    Width =690
                    Height =300
                    FontSize =10
                    TabIndex =31
                    BackColor =255
                    Name ="OnOrder"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Sum(PurchaseOrderLine.QuantityShipped) AS SumOfQuantityShipped F"
                        "ROM Item LEFT JOIN PurchaseOrderLine ON Item.ID=PurchaseOrderLine.ItemID GROUP B"
                        "Y Item.ID; "
                    ColumnWidths ="0;1440;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click on this number to see what Purchase Order(s) have this item."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =540
                            Top =3300
                            Width =780
                            Height =245
                            Name ="On Order_Label"
                            Caption ="On Order"
                            FontName ="Trebuchet MS"
                            EventProcPrefix ="On_Order_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =2925
                    Top =3305
                    Width =675
                    Height =300
                    FontSize =10
                    TabIndex =32
                    BackColor =8453888
                    Name ="OnReserve"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OrderLine.ItemID, Sum(OrderLine.QuantityShipped) AS SumOfQuantityShipped "
                        "FROM OrderLine GROUP BY OrderLine.ItemID; "
                    ColumnWidths ="0;1440;1440"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click on this number to see what Order(s) have this item."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2220
                            Top =3300
                            Width =675
                            Height =285
                            Name ="Label110"
                            Caption ="Reserve"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =11580
                    Top =2580
                    Width =666
                    Height =606
                    TabIndex =33
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
                    UnicodeAccessKey =101

                    LayoutCachedLeft =11580
                    LayoutCachedTop =2580
                    LayoutCachedWidth =12246
                    LayoutCachedHeight =3186
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =12240
                            Top =2580
                            Width =660
                            Height =495
                            Name ="lblDeleteRecord"
                            Caption ="D&elete Item"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =12240
                            LayoutCachedTop =2580
                            LayoutCachedWidth =12900
                            LayoutCachedHeight =3075
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =11580
                    Top =1920
                    Width =666
                    Height =606
                    TabIndex =34
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
                    UnicodeAccessKey =118

                    LayoutCachedLeft =11580
                    LayoutCachedTop =1920
                    LayoutCachedWidth =12246
                    LayoutCachedHeight =2526
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =12255
                            Top =2100
                            Width =555
                            Height =210
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =12255
                            LayoutCachedTop =2100
                            LayoutCachedWidth =12810
                            LayoutCachedHeight =2310
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =11580
                    Top =1260
                    Width =666
                    Height =606
                    TabIndex =35
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
                    UnicodeAccessKey =110

                    LayoutCachedLeft =11580
                    LayoutCachedTop =1260
                    LayoutCachedWidth =12246
                    LayoutCachedHeight =1866
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =12255
                            Top =1380
                            Width =585
                            Height =255
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =12255
                            LayoutCachedTop =1380
                            LayoutCachedWidth =12840
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =120
                    Width =12900
                    Height =3900
                    Name ="boxScreen"
                    LayoutCachedTop =120
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =4020
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =77
                    Left =11580
                    Top =3300
                    Width =666
                    Height =606
                    TabIndex =36
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
                    UnicodeAccessKey =109

                    LayoutCachedLeft =11580
                    LayoutCachedTop =3300
                    LayoutCachedWidth =12246
                    LayoutCachedHeight =3906
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =12292
                            Top =3300
                            Width =510
                            Height =525
                            Name ="Label65"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =12292
                            LayoutCachedTop =3300
                            LayoutCachedWidth =12802
                            LayoutCachedHeight =3825
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =65
                    Left =11580
                    Top =600
                    Width =666
                    Height =606
                    TabIndex =37
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

                    LayoutCachedLeft =11580
                    LayoutCachedTop =600
                    LayoutCachedWidth =12246
                    LayoutCachedHeight =1206
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =12263
                            Top =780
                            Width =705
                            Height =285
                            Name ="Label68"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =12263
                            LayoutCachedTop =780
                            LayoutCachedWidth =12968
                            LayoutCachedHeight =1065
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =82
                    ColumnCount =9
                    ListRows =25
                    ListWidth =12960
                    Left =1020
                    Top =540
                    Width =1080
                    Height =300
                    FontSize =10
                    TabIndex =38
                    BoundColumn =4
                    BackColor =16777164
                    Name ="cmbItemsByVendor"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2880;1440;2880;720;1440;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click to activate."
                    UnicodeAccessKey =114

                    LayoutCachedLeft =1020
                    LayoutCachedTop =540
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =360
                            Top =540
                            Width =615
                            Height =285
                            ForeColor =16711680
                            Name ="Label114"
                            Caption ="Vendo&r"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =360
                            LayoutCachedTop =540
                            LayoutCachedWidth =975
                            LayoutCachedHeight =825
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =247
                    Left =6120
                    Top =1740
                    Width =810
                    Height =300
                    FontSize =10
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Location"
                    ControlSource ="Location"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5250
                            Top =1740
                            Width =810
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label43"
                            Caption ="Location"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =3090
                    Top =2580
                    Width =510
                    Height =300
                    FontSize =10
                    TabIndex =39
                    BackColor =8453888
                    Name ="txtAvailSoldToday"
                    ControlSource ="=[SoldTodayCr1]+[SoldTodayCr2]+[SoldTodayCr3]+[SoldTodayCr4]+[SoldTodayCr5]"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =2160
                            Top =2580
                            Width =855
                            Height =285
                            Name ="Label125"
                            Caption ="Sold Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =3090
                    Top =2940
                    Width =510
                    Height =300
                    FontSize =10
                    TabIndex =40
                    BackColor =8453888
                    Name ="txtAvailTransferred"
                    ControlSource ="TransferredToday"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =2190
                            Top =2940
                            Width =900
                            Height =285
                            Name ="Label131"
                            Caption ="XFR Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =3105
                    Top =2220
                    Width =510
                    Height =300
                    FontSize =10
                    TabIndex =41
                    BackColor =8453888
                    Name ="txtAvailOnTheShelf"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =2100
                            Top =2220
                            Width =975
                            Height =285
                            Name ="Label133"
                            Caption ="On the shelf"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =7065
                    Top =3660
                    Width =250
                    TabIndex =42
                    Name ="PLUList"
                    ControlSource ="PLUList"
                    Tag ="TabStopSkip"
                    ControlTipText ="Displayed in Cash Register's PLU List ?"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =6240
                            Top =3660
                            Width =765
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label51"
                            Caption ="PLU List"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =8040
                    Top =3690
                    TabIndex =43
                    Name ="TPR"
                    ControlSource ="TPR"
                    Tag ="TabStopSkip"
                    ControlTipText ="Temporary Price Reduction Flag - Used for Creating Purchase Orders"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =7560
                            Top =3660
                            Width =360
                            Height =240
                            Name ="Label150"
                            Caption ="TPR"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =945
                    Top =3660
                    Width =540
                    TabIndex =44
                    Name ="AverageDailySales"
                    ControlSource ="AverageDailySales"
                    Format ="#.0"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="Double Click to Recalc ADS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =60
                            Top =3660
                            Width =840
                            Height =285
                            Name ="Label153"
                            Caption ="Avg Dly Sls"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =2925
                    Top =3665
                    Width =675
                    Height =300
                    FontSize =10
                    TabIndex =45
                    BackColor =13434828
                    Name ="OnTransfer"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemOnTransfer"
                    ColumnWidths ="0;1440;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click on this number to see what Transfers(s) have this item."

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =2175
                            Top =3660
                            Width =720
                            Height =285
                            Name ="Label159"
                            Caption ="Transfer"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =5460
                    Top =2220
                    Width =600
                    Height =300
                    FontSize =10
                    TabIndex =46
                    BackColor =255
                    Name ="txtSalePrice"
                    FontName ="Trebuchet MS"
                    Tag ="TabStopSkip"
                    ControlTipText ="Sale Price"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =3720
                            Top =2220
                            Width =840
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label162"
                            Caption ="  SALE Price"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1500
                    Top =3660
                    Width =540
                    TabIndex =47
                    Name ="Combo163"
                    ControlSource ="DaysSupplyGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupAverageDaySupplyLevels.ID, lkupAverageDaySupplyLevels.DaysSupply FROM"
                        " lkupAverageDaySupplyLevels ORDER BY lkupAverageDaySupplyLevels.ID; "
                    ColumnWidths ="0;1440"

                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =4080
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =3120
                    Top =60
                    Width =1740
                    Height =300
                    TabIndex =1
                    BackColor =255
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"
                    Format ="mm/dd/yy hh:nn:ss"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2340
                            Top =60
                            Width =720
                            Height =285
                            Name ="Label53"
                            Caption ="LastSold"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    Left =600
                    Top =60
                    Width =1740
                    Height =300
                    ColumnWidth =3090
                    BackColor =255
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="mm/dd/yy hh:nn:ss"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =60
                            Width =495
                            Height =285
                            Name ="Label55"
                            Caption ="Setup"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =9660
                    Top =60
                    Width =1740
                    Height =300
                    ColumnWidth =2070
                    TabIndex =2
                    BackColor =255
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="mm/dd/yy hh:nn:ss"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8880
                            Top =60
                            Width =720
                            Height =285
                            Name ="Label57"
                            Caption ="Changed"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =8880
                            LayoutCachedTop =60
                            LayoutCachedWidth =9600
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =5700
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =3
                    BackColor =255
                    Name ="SoldTodayCr1"
                    ControlSource ="SoldTodayCr1"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4800
                            Top =60
                            Width =900
                            Height =285
                            Name ="Label35"
                            Caption ="Sold Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =6060
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =4
                    BackColor =255
                    Name ="SoldTodayCr2"
                    ControlSource ="SoldTodayCr2"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =6420
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =5
                    BackColor =255
                    Name ="SoldTodayCr3"
                    ControlSource ="SoldTodayCr3"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =6780
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =6
                    BackColor =255
                    Name ="SoldTodayCr4"
                    ControlSource ="SoldTodayCr4"
                    FontName ="Trebuchet MS"

                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Top =480
                    Width =10740
                    Height =2580
                    TabIndex =7
                    Name ="PurchaseOrdersOnFileSubItemMovement"
                    SourceObject ="Form.PurchaseOrdersOnFileSubItemMovement"
                    LinkChildFields ="Code"
                    LinkMasterFields ="ID"
                    OnEnter ="[Event Procedure]"

                    LayoutCachedTop =480
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =3060
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =8460
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =8
                    BackColor =255
                    Name ="TransferredToday"
                    ControlSource ="TransferredToday"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =7500
                            Top =60
                            Width =900
                            Height =285
                            Name ="Label117"
                            Caption ="XFR Today"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =7500
                            LayoutCachedTop =60
                            LayoutCachedWidth =8400
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =12120
                    Top =60
                    Width =720
                    Height =540
                    FontSize =10
                    BackColor =16763904
                    Name ="lblSaving"
                    Caption ="Saving Record"
                    LayoutCachedLeft =12120
                    LayoutCachedTop =60
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =600
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =11460
                    Top =60
                    Width =720
                    Height =540
                    FontSize =10
                    BackColor =16763904
                    Name ="lblEditing"
                    Caption ="Editing Record"
                    LayoutCachedLeft =11460
                    LayoutCachedTop =60
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10920
                    Top =2175
                    Width =1860
                    Height =420
                    TabIndex =9
                    Name ="cmdPrintLabels"
                    Caption ="Print Labels"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10920
                    LayoutCachedTop =2175
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =2595
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10920
                    Top =2640
                    Width =1860
                    Height =420
                    TabIndex =10
                    Name ="cmdShowDetail"
                    Caption ="Show Detail"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10920
                    LayoutCachedTop =2640
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =3060
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10920
                    Top =3120
                    Width =1860
                    Height =420
                    TabIndex =11
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10920
                    LayoutCachedTop =3120
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =3540
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =3180
                    Width =4440
                    Height =405
                    FontSize =12
                    ForeColor =1279872587
                    Name ="Label170"
                    Caption ="Click Here for recommended Websites:"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =120
                    LayoutCachedTop =3180
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =3585
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10860
                    Top =660
                    Width =666
                    Height =606
                    FontSize =9
                    TabIndex =12
                    Name ="cmdPrintPreview"
                    Caption ="Show All Agents"
                    OnClick ="macKMDSPrintPreview"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000030507050203040ff ,
                        0x203850500000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff504840ff305060ff4078c0ff ,
                        0x304860ff3050605000000000b0a090ffffffffffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffb0a090ffb0a090ff908880ff305070ff3080d0ff4098e0ff ,
                        0x50b0f0ff506870ff00000000b0a090fffffffffffffffffffff8fffff0f0f0ff ,
                        0xd0d8d0ff909890ff706860ff606860ff505860ff407090ff40a0e0ff60c8ffff ,
                        0x7090a0ff7088905000000000b0a090ffffffffffffffffffffffffffe0e0e0ff ,
                        0x909090ffb0a8a0ffd0c0b0ffd0b0a0ff807870ff505850ff6090b0ff7098b0ff ,
                        0x708890800000000000000000b0a090ffffffffffffffffffffffffffb0b0b0ff ,
                        0xc0b8b0fffff0e0ffffe8e0fff0d8c0fff0d0b0ff807870ff6058609070889080 ,
                        0x000000000000000000000000c0a890ffffffffffffffffffffffffffa09890ff ,
                        0xf0e8e0fffff8f0fffff0f0ffffe8e0fff0d8d0ffd0b0a0ff505850e000000000 ,
                        0x000000000000000000000000c0a8a0ffffffffffffffffffffffffffa0a090ff ,
                        0xf0e8e0fffffffffffff8f0fffff0f0ffffe8e0ffe0c0b0ff606060e000000000 ,
                        0x000000000000000000000000c0b0a0ffffffffffffffffffffffffffc0c8c0ff ,
                        0xc0c0c0fffffffffffffffffffff8f0fffff0e0ffb0a090ff6060608000000000 ,
                        0x000000000000000000000000d0b0a0fffffffffffffffffffffffffff0f8ffff ,
                        0xc0b8b0ffc0c0c0fff0e8e0fff0e8e0ffb0b0a0ff707060ff8070601000000000 ,
                        0x000000000000000000000000d0b8a0ffffffffffffffffffffffffffffffffff ,
                        0xf0f8ffffc0c8c0ffa0a090ff909080ff909090ff605040ff0000000000000000 ,
                        0x000000000000000000000000d0b8b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffb0a090ff604830ff604830ff604830ff0000000000000000 ,
                        0x000000000000000000000000d0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a890ffd0c8c0ff604830ffd0b0a0900000000000000000 ,
                        0x000000000000000000000000e0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a8a0ff604830ffd0b0a090000000000000000000000000 ,
                        0x000000000000000000000000e0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b0a0ffd0b0a09000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10860
                    LayoutCachedTop =660
                    LayoutCachedWidth =11526
                    LayoutCachedHeight =1266
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =10860
                            Top =1260
                            Width =1095
                            Height =285
                            Name ="Label61"
                            Caption ="Print Preview"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =10860
                            LayoutCachedTop =1260
                            LayoutCachedWidth =11955
                            LayoutCachedHeight =1545
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12000
                    Top =660
                    Width =666
                    Height =606
                    TabIndex =13
                    ForeColor =0
                    Name ="cmdDataSheet"
                    Caption ="Command59"
                    OnClick ="macKMDSDatasheet"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000806860ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x0000000000000000907860ffd0c8c0ff805840fff0f0f0ffe0d8d0ffd0d0d0ff ,
                        0xd0c8c0ffd0c0c0ffc0b8b0ffc0b0b0ffb0a8a0ffb0a8a0ffa09890ff604830ff ,
                        0x0000000000000000a08070ffd0c8d0ff805840fff0f0f0ffb09080ffa08070ff ,
                        0xe0e0d0ff907050ff906850ffd0c8c0ff906850ff906850ffb09890ff604830ff ,
                        0x0000000000000000a09080ffd0d0d0ff805840fffff8f0fff0f0f0fff0f0f0ff ,
                        0xf0e8e0ffe0e0e0ffe0d8d0ffe0d0d0ffd0c8c0ffd0c0b0ffb0a8a0ff604830ff ,
                        0x0000000000000000b0a090ffd0d0d0ff906040fffff8ffffc0a8a0ffc09890ff ,
                        0xf0f0f0ffa08070ffa07860ffe0d8d0ff906850ff906850ffc0b0b0ff604830ff ,
                        0x0000000000000000c0a8a0ffe0d8e0ff906850fffffffffffff8fffffff8f0ff ,
                        0xf0f0f0fff0f0f0fff0e8e0ffe0e0e0ffe0d8d0ffe0d8d0ffc0b8b0ff604830ff ,
                        0x0000000000000000c0b0a0ffe0e0e0ffa07050ffffffffffc0a8a0ffc0a090ff ,
                        0xfff8f0ffc09890ffb09080fff0e8e0ffa08060ffa07860ffd0c8c0ff604830ff ,
                        0x0000000000000000c0b8a0ffe0e0e0ffa07860ffffffffffffffffffffffffff ,
                        0xfff8fffffff8f0fff0f0f0fff0f0f0fff0e8e0ffe0e0e0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0b8b0fff0e8f0ffb08060ffffffffffc0a890ffc0a090ff ,
                        0xffffffffc0a090ffc0a090fff0f0f0ffb09080ffb09080ffe0d8d0ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffb08870ffffffffffffffffffffffffff ,
                        0xfffffffffffffffffff8fffffff8f0fff0f0f0fff0f0f0ffe0e0e0ff604830ff ,
                        0x0000000000000000d0c0b0ffc09080ffb09070ffb08870ffa08060ffa07860ff ,
                        0x907050ff906850ff906040ff806040ff805840ff805840ff805840ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffc09080fff0f0f0fff0e8f0ffe0e8e0ff ,
                        0xe0e0e0ffe0e0e0ffe0d8e0ffd0d8d0ffd0d0d0ffd0d0d0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ffd0b8b0ffc0b0a0ff ,
                        0xc0b0a0ffc0a8a0ffb0a090ffb09890ffa09080ffa08870ff908070ff907860ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12000
                    LayoutCachedTop =660
                    LayoutCachedWidth =12666
                    LayoutCachedHeight =1266
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =12000
                            Top =1260
                            Width =825
                            Height =285
                            Name ="Label60"
                            Caption ="Datasheet"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =12000
                            LayoutCachedTop =1260
                            LayoutCachedWidth =12825
                            LayoutCachedHeight =1545
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =119
                    TextAlign =2
                    Left =7140
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =14
                    BackColor =255
                    Name ="SoldTodayCr5"
                    ControlSource ="SoldTodayCr5"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =7140
                    LayoutCachedTop =60
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10920
                    Top =1680
                    Width =1860
                    Height =420
                    TabIndex =15
                    Name ="cmdPrintLabelForCurrentItem"
                    Caption ="Print Label for This Item"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10920
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =2100
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9000
                    Top =3120
                    Width =1860
                    Height =420
                    TabIndex =16
                    Name ="cmdPriceTool"
                    Caption ="Pricing Tool"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9000
                    LayoutCachedTop =3120
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =3540
                End
            End
        End
    End
End
CodeBehindForm
' See "Item.cls"
