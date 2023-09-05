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
    Width =11100
    DatasheetFontHeight =11
    ItemSuffix =24
    Right =16790
    Bottom =6730
    RecSrcDt = Begin
        0x19f209e6ed08e640
    End
    RecordSource ="SELECT Table1.[First Name], Table1.[Last Name], Table1.Company, Table1.[Company "
        "Sector], Table1.Job, Table1.Mobile, Table1.Email, Table1.Country, Table1.City, T"
        "able1.[You are interested in], Table1.[How did you hear about us?], Table1.ID FR"
        "OM Table1; "
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin WebBrowser
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7800
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2580
                    Width =2340
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="First Name"
                    ControlSource ="First Name"
                    EventProcPrefix ="First_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2895
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =2580
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label1"
                            Caption ="First Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =2580
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =2895
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =2580
                    Width =2940
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Last Name"
                    ControlSource ="Last Name"
                    EventProcPrefix ="Last_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =5700
                    LayoutCachedTop =2580
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =2895
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4620
                            Top =2580
                            Width =1035
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label2"
                            Caption ="Last Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =4620
                            LayoutCachedTop =2580
                            LayoutCachedWidth =5655
                            LayoutCachedHeight =2895
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =3300
                    Width =2340
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Company"
                    ControlSource ="Company"
                    GridlineColor =10921638

                    LayoutCachedLeft =1560
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =3615
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =3300
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label3"
                            Caption ="Company"
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =3300
                            LayoutCachedWidth =1365
                            LayoutCachedHeight =3615
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =3300
                    Width =2460
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Company Sector"
                    ControlSource ="Company Sector"
                    EventProcPrefix ="Company_Sector"
                    GridlineColor =10921638

                    LayoutCachedLeft =6120
                    LayoutCachedTop =3300
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =3615
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4320
                            Top =3300
                            Width =1575
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label4"
                            Caption ="Company Sector"
                            GridlineColor =10921638
                            LayoutCachedLeft =4320
                            LayoutCachedTop =3300
                            LayoutCachedWidth =5895
                            LayoutCachedHeight =3615
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4860
                    Top =3840
                    Width =3720
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Job"
                    ControlSource ="Job"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =3840
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =4155
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4320
                            Top =3840
                            Width =390
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label5"
                            Caption ="Job"
                            GridlineColor =10921638
                            LayoutCachedLeft =4320
                            LayoutCachedTop =3840
                            LayoutCachedWidth =4710
                            LayoutCachedHeight =4155
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =3840
                    Width =2280
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Mobile"
                    ControlSource ="Mobile"
                    InputMask ="!\\(999\") \"000\\-0000;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =1560
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =4155
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =600
                            Top =3840
                            Width =735
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label6"
                            Caption ="Mobile"
                            GridlineColor =10921638
                            LayoutCachedLeft =600
                            LayoutCachedTop =3840
                            LayoutCachedWidth =1335
                            LayoutCachedHeight =4155
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =4620
                    Width =2400
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Email"
                    ControlSource ="Email"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =4620
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =4935
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4620
                            Width =900
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label7"
                            Caption ="Email"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =4620
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =4935
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6060
                    Top =4620
                    Width =2220
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Country"
                    ControlSource ="Country"
                    GridlineColor =10921638

                    LayoutCachedLeft =6060
                    LayoutCachedTop =4620
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =4935
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5040
                            Top =4620
                            Width =810
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label8"
                            Caption ="Country"
                            GridlineColor =10921638
                            LayoutCachedLeft =5040
                            LayoutCachedTop =4620
                            LayoutCachedWidth =5850
                            LayoutCachedHeight =4935
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1380
                    Top =5220
                    Width =2400
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="City"
                    ControlSource ="City"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =5220
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =5535
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =540
                            Top =5220
                            Width =435
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label9"
                            Caption ="City"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =5220
                            LayoutCachedWidth =975
                            LayoutCachedHeight =5535
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =2820
                    Top =6000
                    Width =3120
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="You are interested in"
                    ControlSource ="You are interested in"
                    RowSourceType ="Value List"
                    RowSource ="\"Poultry\";\"Farming\";\"Aqua Culture\";\"Food Processing & Packaging\""
                    ColumnWidths ="1440"
                    EventProcPrefix ="You_are_interested_in"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2820
                    LayoutCachedTop =6000
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =6315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =6000
                            Width =2220
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label10"
                            Caption ="You are interested in"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =6000
                            LayoutCachedWidth =2520
                            LayoutCachedHeight =6315
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =6720
                    Top =5280
                    Width =2340
                    Height =315
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="How did you hear about us?"
                    ControlSource ="How did you hear about us?"
                    RowSourceType ="Value List"
                    RowSource ="\"LinkedIn\";\"Twitter\";\"Facebook\";\"Ministry\";\"Heard from Colleague\""
                    ColumnWidths ="1440"
                    EventProcPrefix ="How_did_you_hear_about_us_"
                    GridlineColor =10921638

                    LayoutCachedLeft =6720
                    LayoutCachedTop =5280
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =5595
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3960
                            Top =5280
                            Width =2640
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label11"
                            Caption ="How did you hear about us?"
                            GridlineColor =10921638
                            LayoutCachedLeft =3960
                            LayoutCachedTop =5280
                            LayoutCachedWidth =6600
                            LayoutCachedHeight =5595
                        End
                    End
                End
                Begin Image
                    PictureType =2
                    Left =240
                    Top =60
                    Width =8820
                    Height =1260
                    BorderColor =10921638
                    Name ="Image12"
                    Picture ="Logo-2023"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =1320
                    TabIndex =15
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1380
                    Top =1500
                    Width =6240
                    Height =480
                    FontSize =18
                    FontWeight =700
                    BackColor =14277081
                    BorderColor =8355711
                    Name ="Label13"
                    Caption ="SELF REGISTRATION"
                    GridlineColor =10921638
                    LayoutCachedLeft =1380
                    LayoutCachedTop =1500
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =1980
                    BackShade =85.0
                    ForeThemeColorIndex =9
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =540
                    Top =6780
                    Width =3000
                    Height =600
                    TabIndex =11
                    ForeColor =4210752
                    Name ="Print"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =540
                    LayoutCachedTop =6780
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =7380
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2160
                    Width =2460
                    Height =315
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    Format ="General Number"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2475
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1080
                            Top =2160
                            Width =270
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label15"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =1080
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =2475
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5940
                    Top =6600
                    Width =2880
                    Height =720
                    TabIndex =13
                    ForeColor =4210752
                    Name ="New Registration"
                    Caption ="New Registration"
                    EventProcPrefix ="New_Registration"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="5"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"New Registration\" xmlns=\"http://schemas.microsoft.com/offi"
                                "ce/accessservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><A"
                                "ction Name=\"GoToRecord\"><Argume"
                        End
                        Begin
                            Comment ="_AXL:nt Name=\"Record\">New</Argument></Action><ConditionalBlock><If><Condition>"
                                "[MacroError]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argume"
                                "nt Name=\"Message\">=[MacroError].[Description]</Argument></Action></Statements>"
                                "</If></ConditionalB"
                        End
                        Begin
                            Comment ="_AXL:lock></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =5940
                    LayoutCachedTop =6600
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =7320
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin WebBrowser
                    OverlapFlags =85
                    Left =8880
                    Top =1440
                    Width =1680
                    Height =1380
                    AutoActivate =1
                    TabIndex =14
                    BorderColor =16777215
                    Name ="WebBrowser19"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
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
                        0xfffffffffdfffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
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
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000feffffff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000052006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff0100000061f956880a34d011a96b00c0 ,
                        0x4fd705a2000000000000000000000000f09c5f21edbad9010500000000010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000001000000ac000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
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
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff0200000003000000feffffffffffffffffffffffffffffff ,
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
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff38000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x0000000061f956880a34d011a96b00c04fd705a24c0000002a08000083060000 ,
                        0x0000000000000000000000000000000000000000000000004c00000000000000 ,
                        0x0000000001000000e0d057007335cf11ae6908002b2e12620800000000000000 ,
                        0x4c0000000114020000000000c000000000000046800000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000001000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =8880
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =2820
                    ControlSource ="=\"https://chart.googleapis.com/chart?chl=\" & [ID] & \"&chs=100x100&cht=qr\""
                    ScrollBarsVisible =2
                    HyperlinkBinderDescription ="1|=\"https://chart.googleapis.com/chart?chl=\" & [ID] & \"&chs=100x100&cht=qr\""
                    BorderShade =100.0
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Print_Click()

    
    
    Dim a As String
a = "ID =" & Me.ID
Me.Dirty = False
DoCmd.OpenReport "BadgePrint", acViewPreview, wherecondition:=a
DoCmd.PrintOut PrintRange:=acPages
End Sub
