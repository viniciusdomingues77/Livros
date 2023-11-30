Attribute VB_Name = "Relatorio"
Option Explicit
Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type
Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hDC As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Public Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Public Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public OFName As OPENFILENAME
Dim CustomColors() As Byte
