Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13860
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =-2280
    Top =765
    Right =9900
    Bottom =7395
    RecSrcDt = Begin
        0x8af1119ae898e440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
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
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
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
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
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
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
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
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri"
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
            Height =6960
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Left =225
                    Top =240
                    Width =13500
                    Height =6480
                    Name ="TabCtl6"

                    LayoutCachedLeft =225
                    LayoutCachedTop =240
                    LayoutCachedWidth =13725
                    LayoutCachedHeight =6720
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =360
                            Top =720
                            Width =13235
                            Height =5865
                            Name ="Page7"
                            Caption ="Compare Purchase Orders"
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
                            LayoutCachedLeft =360
                            LayoutCachedTop =720
                            LayoutCachedWidth =13595
                            LayoutCachedHeight =6585
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =365
                                    Top =1265
                                    Width =6360
                                    Height =4545
                                    FontSize =10
                                    BackColor =13434828
                                    BorderColor =0
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
                                    FontName ="Tahoma"
                                    AllowValueListEdits =0
                                    InheritValueList =0
                                    LayoutCachedLeft =365
                                    LayoutCachedTop =1265
                                    LayoutCachedWidth =6725
                                    LayoutCachedHeight =5810
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =360
                                            Top =780
                                            Width =4980
                                            Height =240
                                            FontSize =8
                                            Name ="Outstanding Orders_Label"
                                            Caption ="To compare purchase orders. First select the ORIGINAL order below"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="Outstanding_Orders_Label"
                                            LayoutCachedLeft =360
                                            LayoutCachedTop =780
                                            LayoutCachedWidth =5340
                                            LayoutCachedHeight =1020
                                        End
                                    End
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    Visible = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =7235
                                    Top =1280
                                    Width =6360
                                    Height =4545
                                    FontSize =10
                                    TabIndex =1
                                    BackColor =15195103
                                    BorderColor =0
                                    Name ="lstInvoices_2"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT PurchaseOrder.ID, Vendor.Name, Format([OrderDate],\"mm/dd\") AS [Date], P"
                                        "urchaseOrder.OrderCount AS InvNumber, qryPurchaseOrderLineTotals.Quantity, qryPu"
                                        "rchaseOrderLineTotals.Dollars FROM (PurchaseOrder INNER JOIN Vendor ON PurchaseO"
                                        "rder.VendorID = Vendor.ID) INNER JOIN qryPurchaseOrderLineTotals ON PurchaseOrde"
                                        "r.ID = qryPurchaseOrderLineTotals.PurchaseOrderID WHERE (((PurchaseOrder.ID)<>[F"
                                        "orms]![PurchaseOrderCompare]![txtPurchaseOrderID_1])) ORDER BY Vendor.Name, Form"
                                        "at([OrderDate],\"mm/dd\");"
                                    ColumnWidths ="0;2160;720;720;1080;1080"
                                    BeforeUpdate ="[Event Procedure]"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Tahoma"
                                    AllowValueListEdits =0
                                    InheritValueList =0
                                    LayoutCachedLeft =7235
                                    LayoutCachedTop =1280
                                    LayoutCachedWidth =13595
                                    LayoutCachedHeight =5825
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =7235
                                            Top =800
                                            Width =4650
                                            Height =435
                                            FontSize =8
                                            Name ="Label1"
                                            Caption ="Then select the second (scanned) order below"
                                            FontName ="Tahoma"
                                            LayoutCachedLeft =7235
                                            LayoutCachedTop =800
                                            LayoutCachedWidth =11885
                                            LayoutCachedHeight =1235
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =3000
                                    Top =6000
                                    Height =315
                                    TabIndex =2
                                    Name ="txtPurchaseOrderID_1"
                                    DefaultValue ="=0"
                                    LayoutCachedLeft =3000
                                    LayoutCachedTop =6000
                                    LayoutCachedWidth =4440
                                    LayoutCachedHeight =6315
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =780
                                            Top =6000
                                            Width =1950
                                            Height =315
                                            Name ="Label3"
                                            Caption ="First Purchase Order"
                                            LayoutCachedLeft =780
                                            LayoutCachedTop =6000
                                            LayoutCachedWidth =2730
                                            LayoutCachedHeight =6315
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =9660
                                    Top =6000
                                    Height =315
                                    TabIndex =3
                                    Name ="txtPurchaseOrderID_2"
                                    LayoutCachedLeft =9660
                                    LayoutCachedTop =6000
                                    LayoutCachedWidth =11100
                                    LayoutCachedHeight =6315
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =7440
                                            Top =6000
                                            Width =1950
                                            Height =315
                                            Name ="Label5"
                                            Caption ="2nd  Purchase Order"
                                            LayoutCachedLeft =7440
                                            LayoutCachedTop =6000
                                            LayoutCachedWidth =9390
                                            LayoutCachedHeight =6315
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =223
                    AccessKey =67
                    Left =12360
                    Top =5940
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =1
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
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12360
                    LayoutCachedTop =5940
                    LayoutCachedWidth =13026
                    LayoutCachedHeight =6546
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =13080
                            Top =5940
                            Width =585
                            Height =405
                            FontSize =8
                            Name ="Label65"
                            Caption ="&Close Form"
                            FontName ="Tahoma"
                            LayoutCachedLeft =13080
                            LayoutCachedTop =5940
                            LayoutCachedWidth =13665
                            LayoutCachedHeight =6345
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =13800
                    Height =6900
                    Name ="boxScreen"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =6960
                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrderCompare.cls"
