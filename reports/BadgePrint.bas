Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2820
    DatasheetFontHeight =11
    ItemSuffix =24
    Filter ="ID =-923594478"
    RecSrcDt = Begin
        0x77953e6bd508e640
    End
    RecordSource ="Table1"
    Caption ="Badge Print"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
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
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1320
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    ScrollBars =2
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =17
                    IMESentenceMode =3
                    Top =840
                    Width =2160
                    Height =360
                    FontSize =36
                    BorderColor =10921638
                    ForeColor =-2147483615
                    Name ="ID"
                    ControlSource ="ID"
                    Format ="General Number"
                    FontName ="Code39"
                    GridlineColor =10921638

                    LayoutCachedTop =840
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1200
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Width =1380
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="First Name"
                    ControlSource ="First Name"
                    FontName ="Almarai"
                    EventProcPrefix ="First_Name"
                    GridlineColor =10921638

                    LayoutCachedWidth =1380
                    LayoutCachedHeight =300
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1440
                    Width =1260
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Last Name"
                    ControlSource ="Last Name"
                    FontName ="Almarai"
                    EventProcPrefix ="Last_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =300
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Top =360
                    Width =2760
                    FontSize =10
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Job"
                    ControlSource ="Job"
                    FontName ="Almarai"
                    GridlineColor =10921638

                    LayoutCachedTop =360
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =600
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Top =600
                    Width =2760
                    FontSize =10
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Company"
                    ControlSource ="Company"
                    FontName ="Almarai"
                    GridlineColor =10921638

                    LayoutCachedTop =600
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =840
                    ThemeFontIndex =-1
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
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
     ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub btnGenerateQR_Click()
   ' Declare variables
    Dim apiUrl As String
    Dim qrData As String
    Dim savePath As String
    Dim result As Long

    ' Get QR data from textbox
    qrData = Me.ID.Value

    ' Construct API URL
    apiUrl = "https://api.qrserver.com/v1/create-qr-code/?data=" & qrData & "&size=200x200"

    ' Specify save path for BMP file in the same directory as the Access database
    savePath = Application.CurrentProject.Path & "\qr_code.bmp"

    ' Download QR code image as BMP file
    result = URLDownloadToFile(0, apiUrl, savePath, 0, 0)

    ' Check if download was successful
    If result = 0 Then
        ' Display the downloaded image in the image control
        Me.imgQRCode.Picture = savePath
    Else
        MsgBox "Failed to download QR code image.", vbExclamation
    End If
End Sub
