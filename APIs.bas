Attribute VB_Name = "Module5"
'This module contains API declarations
'Last Modified 01/02/2002 nt
'-----------------------------------------------
'windows message constants
Public Const WM_NCLBUTTONDOWN = &HA1
'Public Const WM_USER = &H400

'Public Const CC_ANYCOLOR = &H100
'Public Const CC_CHORD = 4
'Public Const CC_CIRCLES = 1
'Public Const CC_ELLIPSES = 8
'Public Const CC_ENABLEHOOK = &H10
'Public Const CC_ENABLETEMPLATE = &H20
'Public Const CC_ENABLETEMPLATEHANDLE = &H40
'Public Const CC_FULLOPEN = &H2
'Public Const CC_INTERIORS = 128
'Public Const CC_NONE = 0
'Public Const CC_PIE = 2
'Public Const CC_PREVENTFULLOPEN = &H4
'Public Const CC_RGBINIT = &H1
'Public Const CC_ROUNDRECT = 256
'Public Const CC_SHOWHELP = &H8
'Public Const CC_SOLIDCOLOR = &H80
'Public Const CC_STYLED = 32
'Public Const CC_WIDE = 16
'Public Const CC_WIDESTYLED = 64

'Public Const CF_ANSIONLY = &H400&
'Public Const CF_APPLY = &H200&
'Public Const CF_BITMAP = 2
'Public Const CF_DIB = 8
'Public Const CF_DIF = 5
'Public Const CF_DSPBITMAP = &H82
'Public Const CF_DSPENHMETAFILE = &H8E
'Public Const CF_DSPMETAFILEPICT = &H83
'Public Const CF_DSPTEXT = &H81
'Public Const CF_EFFECTS = &H100&
'Public Const CF_ENABLEHOOK = &H8&
'Public Const CF_ENABLETEMPLATE = &H10&
'Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_ENHMETAFILE = 14
'Public Const CF_FIXEDPITCHONLY = &H4000&
'Public Const CF_FORCEFONTEXIST = &H10000
'Public Const CF_GDIOBJFIRST = &H300
'Public Const CF_GDIOBJLAST = &H3FF
'Public Const CF_INITTOLOGFONTSTRUCT = &H40&
'Public Const CF_LIMITSIZE = &H2000&
Public Const CF_METAFILEPICT = 3
'Public Const CF_NOFACESEL = &H80000
'Public Const CF_NOSCRIPTSEL = &H800000
'Public Const CF_NOSIMULATIONS = &H1000&
'Public Const CF_NOSIZESEL = &H200000
'Public Const CF_NOSTYLESEL = &H100000
'Public Const CF_NOVECTORFONTS = &H800&
'Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
'Public Const CF_NOVERTFONTS = &H1000000
'Public Const CF_OEMTEXT = 7
'Public Const CF_OWNERDISPLAY = &H80
'Public Const CF_PALETTE = 9
'Public Const CF_PENDATA = 10
'Public Const CF_PRINTERFONTS = &H2
'Public Const CF_PRIVATEFIRST = &H200
'Public Const CF_PRIVATELAST = &H2FF
'Public Const CF_RIFF = 11
'Public Const CF_SCALABLEONLY = &H20000
'Public Const CF_SCREENFONTS = &H1
'Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Public Const CF_SCRIPTSONLY = CF_ANSIONLY
'Public Const CF_SELECTSCRIPT = &H400000
'Public Const CF_SHOWHELP = &H4&
'Public Const CF_SYLK = 4
'Public Const CF_TEXT = 1
'Public Const CF_TIFF = 6
'Public Const CF_TTONLY = &H40000
'Public Const CF_UNICODETEXT = 13
'Public Const CF_USESTYLE = &H80&
'Public Const CF_WAVE = 12
'Public Const CF_WYSIWYG = &H8000

'Public Const CP_ACP = 0
'Public Const CP_NONE = 0
'Public Const CP_OEMCP = 1
'Public Const CP_RECTANGLE = 1
'Public Const CP_REGION = 2
'Public Const CP_WINANSI = 1004
'Public Const CP_WINUNICODE = 1200

'DEVMODE structure constants
'Public Const DM_COLLATE As Long = &H8000
'Public Const DM_COLOR = &H800&
'Public Const DM_COPIES = &H100&
'Public Const DM_COPY = 2
'Public Const DM_DEFAULTSOURCE = &H200&
'Public Const DM_DITHERTYPE = &H10000000
'Public Const DM_DUPLEX = &H1000&
'Public Const DM_FORMNAME As Long = &H10000
'Public Const DM_GRAYSCALE = &H1
'Public Const DM_ICMINTENT = &H4000000
'Public Const DM_ICMMETHOD = &H2000000
'Public Const DM_INTERLACED = &H2
'Public Const DM_MEDIATYPE = &H8000000
'Public Const DM_MODIFY = 8
Public Const DM_ORIENTATION = &H1&
'Public Const DM_PAPERLENGTH = &H4&
'Public Const DM_PAPERSIZE = &H2&
'Public Const DM_PAPERWIDTH = &H8&
'Public Const DM_PRINTQUALITY = &H400&
'Public Const DM_PROMPT = 4
'Public Const DM_RESERVED1 = &H800000
'Public Const DM_RESERVED2 = &H1000000
'Public Const DM_SCALE = &H10&
'Public Const DM_SPECVERSION = &H320
'Public Const DM_TTOPTION = &H4000&
'Public Const DM_UPDATE = 1
'Public Const DM_YRESOLUTION = &H2000&
'Public Const DMORIENT_LANDSCAPE = 2
Public Const DMORIENT_PORTRAIT = 1
'Public Const DM_IN_BUFFER = DM_MODIFY
'Public Const DM_IN_PROMPT = DM_PROMPT
'Public Const DM_OUT_DEFAULT = DM_UPDATE
'Public Const DM_GETDEFID = WM_USER + 0
'Public Const DM_SETDEFID = WM_USER + 1
'Public Const DM_OUT_BUFFER = DM_COPY

'Public Const DT_BOTTOM = &H8
'Public Const DT_CALCRECT = &H400
'Public Const DT_CENTER = &H1
'Public Const DT_CHARSTREAM = 4
'Public Const DT_DISPFILE = 6
'Public Const DT_EXPANDTABS = &H40
'Public Const DT_EXTERNALLEADING = &H200
'Public Const DT_INTERNAL = &H1000
'Public Const DT_LEFT = &H0
'Public Const DT_METAFILE = 5
'Public Const DT_NOCLIP = &H100
'Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0
Public Const DT_RASCAMERA = 3
Public Const DT_RASDISPLAY = 1
Public Const DT_RASPRINTER = 2
'Public Const DT_RIGHT = &H2
'Public Const DT_DoubleLINE = &H20
'Public Const DT_TABSTOP = &H80
'Public Const DT_TOP = &H0
'Public Const DT_VCENTER = &H4
'Public Const DT_WORDBREAK = &H10
'global memory constants
'Public Const GMEM_DDESHARE = &H2000
'Public Const GMEM_DISCARDABLE = &H100
'Public Const GMEM_DISCARDED = &H4000
'Public Const GMEM_FIXED = &H0
'Public Const GMEM_INVALID_HANDLE = &H8000
'Public Const GMEM_LOCKCOUNT = &HFF
'Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
'Public Const GMEM_NOCOMPACT = &H10
'Public Const GMEM_NODISCARD = &H20
'Public Const GMEM_NOT_BANKED = &H1000
'Public Const GMEM_LOWER = GMEM_NOT_BANKED
'Public Const GMEM_NOTIFY = &H4000
'Public Const GMEM_SHARE = &H2000
'Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
'Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

'registry keys
'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const HKEY_CURRENT_CONFIG = &H80000005
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const HKEY_DYN_DATA = &H80000006
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_PERFORMANCE_DATA = &H80000004
'Public Const HKEY_USERS = &H80000003

'function call results constants
'Public Const ERROR_SUCCESS = 0&
'Public Const ERROR_FILE_NOT_FOUND = 2&
'Public Const ERROR_PATH_NOT_FOUND = 3&
'Public Const ERROR_BAD_FORMAT = 11&

''''WinHelp constants and function
'''Public Const HELP_COMMAND = &H102&
'''Public Const HELP_CONTENTS = &H3&
'''Public Const HELP_CONTEXT = &H1
'''Public Const HELP_CONTEXTPOPUP = &H8&
'''Public Const HELP_FINDER = &HB& 'this one is not listed
'''Public Const HELP_FORCEFILE = &H9&
'''Public Const HELP_HELPONHELP = &H4
'''Public Const HELP_INDEX = &H3
'''Public Const HELP_KEY = &H101
'''Public Const HELP_MULTIKEY = &H201&
'''Public Const HELP_PARTIALKEY = &H105&
'''Public Const HELP_QUIT = &H2
'''Public Const HELP_SETCONTENTS = &H5&
'''Public Const HELP_SETINDEX = &H5
'''Public Const HELP_SETWINPOS = &H203&

'Public Const HS_API_MAX = 25
'Public Const HS_BDIAGONAL = 3
'Public Const HS_BDIAGONAL1 = 7
'Public Const HS_CROSS = 4
'Public Const HS_DENSE1 = 9
'Public Const HS_DENSE2 = 10
'Public Const HS_DENSE3 = 11
'Public Const HS_DENSE4 = 12
'Public Const HS_DENSE5 = 13
'Public Const HS_DENSE6 = 14
'Public Const HS_DENSE7 = 15
'Public Const HS_DENSE8 = 16
'Public Const HS_DIAGCROSS = 5
'Public Const HS_DITHEREDBKCLR = 24
'Public Const HS_DITHEREDCLR = 20
'Public Const HS_DITHEREDTEXTCLR = 22
'Public Const HS_FDIAGONAL = 2
'Public Const HS_FDIAGONAL1 = 6
'Public Const HS_HALFTONE = 18
'Public Const HS_HORIZONTAL = 0
'Public Const HS_NOSHADE = 17
'Public Const HS_SOLID = 8
'Public Const HS_SOLIDBKCLR = 23
'Public Const HS_SOLIDCLR = 19
'Public Const HS_SOLIDTEXTCLR = 21
'Public Const HS_VERTICAL = 1

'Public Const HWND_BOTTOM = 1
'Public Const HWND_BROADCAST = &HFFFF&
'Public Const HWND_DESKTOP = 0
'Public Const HWND_NOTOPMOST = -2
'Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

'Public Const LC_INTERIORS = 128
'Public Const LC_MARKER = 4
'Public Const LC_NONE = 0
'Public Const LC_POLYLINE = 2
'Public Const LC_POLYMARKER = 8
'Public Const LC_STYLED = 32
'Public Const LC_WIDE = 16
'Public Const LC_WIDESTYLED = 64

'logical font constants
Public Const LF_FACESIZE = 32
'Public Const LF_FULLFACESIZE = 64

'mapping modes constants
Public Const MM_ANISOTROPIC = 8
'Public Const MM_HIENGLISH = 5
Public Const MM_HIMETRIC = 3
'Public Const MM_ISOTROPIC = 7
'Public Const MM_TEXT = 1
'Public Const MM_TWIPS = 6

'OPENFILENAME structure flags constants (used with Open & Save dialogs)
'Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
'Public Const OFN_ENABLEHOOK = &H20
'Public Const OFN_ENABLETEMPLATE = &H40
'Public Const OFN_ENABLETEMPLATEHANDLE = &H80
'Public Const OFN_EXPLORER = &H80000
'Public Const OFN_EXTENSIONDIFFERENT = &H400
'Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
'Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
'Public Const OFN_NOLONGNAMES = &H40000
'Public Const OFN_NONETWORKBUTTON = &H20000
'Public Const OFN_NOREADONLYRETURN = &H8000
'Public Const OFN_NOTESTFILECREATE = &H10000
'Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
'Public Const OFN_PATHMUSTEXIST = &H800
'Public Const OFN_READONLY = &H1
'Public Const OFN_SHAREAWARE = &H4000
'Public Const OFN_SHAREFALLTHROUGH = 2
'Public Const OFN_SHARENOWARN = 1
'Public Const OFN_SHAREWARN = 0
'Public Const OFN_SHOWHELP = &H10
'Public Const OFS_MAXPATHNAME = 128

'nonstandard constants to save long statements
Public Const OFS_OPENFILE_FLAGS = OFN_LONGNAMES Or _
                                  OFN_CREATEPROMPT Or _
                                  OFN_NODEREFERENCELINKS

Public Const OFS_SAVEFILE_FLAGS = OFN_LONGNAMES Or _
                                  OFN_OVERWRITEPROMPT Or _
                                  OFN_HIDEREADONLY


'Public Const OBJ_BITMAP = 7
'Public Const OBJ_BRUSH = 2
'Public Const OBJ_DC = 3
'Public Const OBJ_ENHMETADC = 12
'Public Const OBJ_ENHMETAFILE = 13
'Public Const OBJ_EXTPEN = 11
'Public Const OBJ_FONT = 6
'Public Const OBJ_MEMDC = 10
'Public Const OBJ_METADC = 4
'Public Const OBJ_METAFILE = 9
'Public Const OBJ_PAL = 5
'Public Const OBJ_PEN = 1
'Public Const OBJ_REGION = 8

'Public Const PC_EXPLICIT = &H2
'Public Const PC_INTERIORS = 128
'Public Const PC_NOCOLLAPSE = &H4
'Public Const PC_NONE = 0
'Public Const PC_POLYGON = 1
'Public Const PC_RECTANGLE = 2
'Public Const PC_RESERVED = &H1
'Public Const PC_SCANLINE = 8
'Public Const PC_STYLED = 32
'Public Const PC_TRAPEZOID = 4
'Public Const PC_WIDE = 16
'Public Const PC_WIDESTYLED = 64
'Public Const PC_WINDPOLYGON = 4

'printer common dialog constants
'Public Const PD_ALLPAGES = &H0
'Public Const PD_COLLATE = &H10
'Public Const PD_DISABLEPRINTTOFILE = &H80000
'Public Const PD_ENABLEPRINTHOOK = &H1000
'Public Const PD_ENABLEPRINTTEMPLATE = &H4000
'Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
'Public Const PD_ENABLESETUPHOOK = &H2000
'Public Const PD_ENABLESETUPTEMPLATE = &H8000
'Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
'Public Const PD_HIDEPRINTTOFILE = &H100000
'Public Const PD_NONETWORKBUTTON = &H200000
'Public Const PD_NOPAGENUMS = &H8
'Public Const PD_NOSELECTION = &H4
'Public Const PD_NOWARNING = &H80
'Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
'Public Const PD_PRINTTOFILE = &H20
'Public Const PD_RETURNDC = &H100
'Public Const PD_RETURNDEFAULT = &H400
'Public Const PD_RETURNIC = &H200
'Public Const PD_SELECTION = &H1
'Public Const PD_SHOWHELP = &H800
'Public Const PD_USEDEVMODECOPIES = &H40000
'Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

'Public Const PS_ALTERNATE = 8
'Public Const PS_COSMETIC = &H0
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_DOT = 2
'Public Const PS_ENDCAP_FLAT = &H200
'Public Const PS_ENDCAP_MASK = &HF00
'Public Const PS_ENDCAP_ROUND = &H0
'Public Const PS_ENDCAP_SQUARE = &H100
'Public Const PS_GEOMETRIC = &H10000
'Public Const PS_INSIDEFRAME = 6
'Public Const PS_JOIN_BEVEL = &H1000
'Public Const PS_JOIN_MASK = &HF000
'Public Const PS_JOIN_MITER = &H2000
'Public Const PS_JOIN_ROUND = &H0
'Public Const PS_NULL = 5
Public Const PS_SOLID = 0
'Public Const PS_STYLE_MASK = &HF
'Public Const PS_TYPE_MASK = &HF0000
'Public Const PS_USERSTYLE = 7

'Public Const R2_BLACK = 1
Public Const R2_COPYPEN = 13
'Public Const R2_LAST = 16
'Public Const R2_MASKNOTPEN = 3
Public Const R2_MASKPEN = 9
'Public Const R2_MASKPENNOT = 5
'Public Const R2_MERGENOTPEN = 12
'Public Const R2_MERGEPEN = 15
'Public Const R2_MERGEPENNOT = 14
'Public Const R2_NOP = 11
'Public Const R2_NOT = 6
'Public Const R2_NOTCOPYPEN = 4
'Public Const R2_NOTMASKPEN = 8
'Public Const R2_NOTMERGEPEN = 2
'Public Const R2_NOTXORPEN = 10
'Public Const R2_WHITE = 16
'Public Const R2_XORPEN = 7

Public Const RC_BANDING = 2
Public Const RC_BIGFONT = &H400
Public Const RC_BITBLT = 1
Public Const RC_BITMAP64 = 8
'Public Const RC_DEVBITS = &H8000
Public Const RC_DI_BITMAP = &H80
Public Const RC_DIBTODEV = &H200
Public Const RC_FLOODFILL = &H1000
'Public Const RC_GDI20_OUTPUT = &H10
'Public Const RC_GDI20_STATE = &H20
'Public Const RC_NONE = 0
'Public Const RC_OP_DX_OUTPUT = &H4000
'Public Const RC_PALETTE = &H100
'Public Const RC_SAVEBITMAP = &H40
Public Const RC_SCALING = 4
Public Const RC_STRETCHBLT = &H800
Public Const RC_STRETCHDIB = &H2000

'registry key types
'Public Const REG_BINARY = 3
'Public Const REG_DWORD = 4
'Public Const REG_DWORD_BIG_ENDIAN = 5
'Public Const REG_DWORD_LITTLE_ENDIAN = 4
'Public Const REG_EXPAND_SZ = 2
'Public Const REG_LINK = 6
'Public Const REG_MULTI_SZ = 7
'Public Const REG_NONE = 0
'Public Const REG_SZ = 1

'registry security constants
'Public Const READ_CONTROL = &H20000
'Public Const KEY_QUERY_VALUE = &H1
'Public Const KEY_SET_VALUE = &H2
'Public Const KEY_CREATE_SUB_KEY = &H4
'Public Const KEY_ENUMERATE_SUB_KEYS = &H8
'Public Const KEY_NOTIFY = &H10
'Public Const KEY_CREATE_LINK = &H20
'Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
'                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
'                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

'Blting raster operations
'Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
'Public Const SRCERASE = &H440328
'Public Const SRCINVERT = &H660046
'Public Const SRCPAINT = &HEE0086
'Public Const NOTSRCCOPY = &H330008
'Public Const NOTSRCERASE = &H1100A6
'Public Const MERGECOPY = &HC000CA
'Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
'Public Const PATINVERT = &H5A0049
'Public Const PATPAINT = &HFB0A09
'Public Const DSTINVERT = &H550009
'Public Const BLACKNESS = &H42
'Public Const WHITENESS = &HFF0062

'SystemMetrics constants
'Public Const SM_CXBORDER = 5
'Public Const SM_CYBORDER = 6
'Public Const SM_CYCAPTION = 4
'Public Const SM_CYDLGFRAME = 8
'Public Const SM_CYFRAME = 33
Public Const SM_CYMENU = 15
Public Const SM_CXSCREEN = 0
'Public Const SM_CYSCREEN = 1
'Public Const SM_CXHSCROLL = 21
'Public Const SM_CXVSCROLL = 2
'Public Const SM_CYHSCROLL = 3
'Public Const SM_CYVSCROLL = 20

'set window position constants
'Public Const SWP_FRAMECHANGED = &H20
'Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_NOCOPYBITS = &H100
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOOWNERZORDER = &H200
'Public Const SWP_NOREDRAW = &H8
'Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOZORDER = &H4
'Public Const SWP_SHOWWINDOW = &H40

'Show window mode constants
Public Const SW_HIDE = 0
'Public Const SW_SHOW = 5
'Public Const SW_SHOWDEFAULT = 10
'Public Const SW_SHOWMAXIMIZED = 3
'Public Const SW_SHOWMINIMIZED = 2
'Public Const SW_SHOWNORMAL = 1
'Public Const SW_SHOWNOACTIVATE = 4
'Public Const SW_SHOWMINNOACTIVE = 7
'Public Const SW_SHOWNA = 8

'text align constants
'Public Const TA_BASELINE = 24
'Public Const TA_BOTTOM = 8
'Public Const TA_CENTER = 6
'Public Const TA_LEFT = 0
'Public Const TA_NOUPDATECP = 0
'Public Const TA_RIGHT = 2
'Public Const TA_TOP = 0
'Public Const TA_UPDATECP = 1
'Public Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)


'Public Const TC_CP_STROKE = &H4
'Public Const TC_CR_90 = &H8
'Public Const TC_CR_ANY = &H10
'Public Const TC_EA_DOUBLE = &H200
'Public Const TC_GP_TRAP = 2
'Public Const TC_HARDERR = 1
'Public Const TC_IA_ABLE = &H400
'Public Const TC_NORMAL = 0
'Public Const TC_OP_CHARACTER = &H1
'Public Const TC_OP_STROKE = &H2
'Public Const TC_RA_ABLE = &H2000
'Public Const TC_RESERVED = &H8000
'Public Const TC_SA_CONTIN = &H100
'Public Const TC_SA_DOUBLE = &H40
'Public Const TC_SA_INTEGER = &H80
'Public Const TC_SCROLLBLT = &H10000
'Public Const TC_SF_X_YINDEP = &H20
'Public Const TC_SIGNAL = 3
'Public Const TC_SO_ABLE = &H1000
'Public Const TC_UA_ABLE = &H800
'Public Const TC_VA_ABLE = &H4000
'OS version constants
'Public Const VER_PLATFORM_WIN32_NT = 2
'Public Const VER_PLATFORM_WIN32_WINDOWS = 1
'Public Const VER_PLATFORM_WIN32s = 0
'Miscellaneous constants
Public Const ASPECTX = 40
Public Const ASPECTXY = 44
Public Const ASPECTY = 42
Public Const BITSPIXEL = 12
'Public Const BLACK_PEN = 7
'Public Const CLIPCAPS = 36
'Public Const COLORRES = 108
'Public Const CURVECAPS = 28
'Public Const DEFAULT_GUI_FONT = 17
'Public Const DRIVERVERSION = 0
'Public Const FLOODFILLBORDER = 0
'Public Const FLOODFILLSURFACE = 1
Public Const HORZRES = 8
Public Const HORZSIZE = 4
Public Const HTCAPTION = 2
'Public Const LINECAPS = 30
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const NULL_BRUSH = 5
Public Const NULL_PEN = 8
'Public Const NUMBRUSHES = 16
'Public Const NUMPENS = 18
'Public Const NUMMARKERS = 20
'Public Const NUMFONTS = 22
Public Const NUMCOLORS = 24
'Public Const NUMRESERVED = 106
Public Const OPAQUE = 2
'Public Const PDEVICESIZE = 26
'Public Const PHYSICALWIDTH = 110
'Public Const PHYSICALHEIGHT = 111
'Public Const PHYSICALOFFSETX = 112
'Public Const PHYSICALOFFSETY = 113
Public Const PLANES = 14
'Public Const POLYGONALCAPS = 32
Public Const RASTERCAPS = 38
'Public Const SIZEPALETTE = 104
Public Const SYSTEM_FONT = 13
Public Const TECHNOLOGY = 2
'Public Const TEXTCAPS = 34
Public Const TRANSPARENT = 1
Public Const VERTRES = 10
Public Const VERTSIZE = 6
'Public Const WHITE_PEN = 6

'Public Const APINULL = 0&   'not API constant
'next two constants used during printing of the Rich Text Box
'Public Const EM_FORMATRANGE As Long = WM_USER + 57
'Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72

'constants used by Browse For Folder dialog
Public Const BIF_RETURNONLYFSDIRS = &H1
'Public Const BIF_DONTGOBELOWDOMAIN = &H2
'Public Const BIF_STATUSTEXT = &H4
'Public Const BIF_RETURNFSANCESTORS = &H8
'Public Const BIF_BROWSEFORCOMPUTER = &H1000
'Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const MAX_PATH = 260

'Structures
'Private Type OSVERSIONINFO
'        dwOSVersionInfoSize As Long
'        dwMajorVersion As Long
'        dwMinorVersion As Long
'        dwBuildNumber As Long
'        dwPlatformId As Long
'        szCSDVersion As String * 128      '  Maintenance string for PSS usage
'End Type

Public Type Size
        cx As Long
        cy As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Private Type RECTS  'necessary for standard metafiles
'        Left As Integer
'        Top As Integer
'        Right As Integer
'        Bottom As Integer
'End Type

Public Type METAFILEPICT
        mm As Long
        xExt As Long
        yExt As Long
        hMF As Long
End Type

'Private Type METAHEADER
'        mtType As Integer
'        mtHeaderSize As Integer
'        mtVersion As Integer
'        mtSize As Long
'        mtNoObjects As Integer
'        mtMaxRecord As Long
'        mtNoParameters As Integer
'End Type

'Private Type METARECORD
'        rdSize As Long
'        rdFunction As Integer
'        rdParm(1) As Integer
'End Type

'Private Type ENHMETAHEADER
'        iType As Long
'        nSize As Long
'        rclBounds As Rect
'        rclFrame As Rect
'        dSignature As Long
'        nVersion As Long
'        nBytes As Long
'        nRecords As Long
'        nHandles As Integer
'        sReserved As Integer
'        nDescription As Long
'        offDescription As Long
'        nPalEntries As Long
'        szlDevice As Size
'        szlMillimeters As Size
'End Type

'Private Type ENHMETARECORD
'        iType As Long
'        nSize As Long
'        dParm(1) As Long
'End Type

'Private Type PALETTEENTRY
'        peRed As Byte
'        peGreen As Byte
'        peBlue As Byte
'        peFlags As Byte
'End Type

'Private Type HANDLETABLE
'        objectHandle(1) As Long
'End Type

Public Type LOGFONT
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
        lfFaceName(LF_FACESIZE) As Byte
End Type

'nexttwo types are used to print content of the Rich Text Box
'Private Type CharRange
'    cpMin As Long   'first char in range - 0 for start of doc
'    cpMax As Long   'last char in range - -1 for end of doc
'End Type

'Private Type FormatRange
'    hDC As Long         'dc to draw on
'    hdcTarget As Long   'tzrget dc to determine text formatting
'    rc As Rect          'region of dc to draw to
'    rcpage As Rect      'region of the entire DC - page size
'    chrg As CharRange   'range of text to draw
'End Type

'used with GetOpenFileName function
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

Public Type DEVMODE
        dmDeviceName As String * 32
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
        dmFormName As String * 32
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        extra As String * 100
End Type

Public Type PrintDlg
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

'structure used to Browse For Folders
Public Type BROWSEINFO
        hwndOwner As Long
        pidlRoot As Long
        lpszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
End Type

'structure used for Color Dialog
Public Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'functions
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Rect) As Long
Public Declare Function ClipCursorByNum Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal ByteCount As Long)
Public Declare Function CopyMetaFile Lib "gdi32" Alias "CopyMetaFileA" (ByVal hMF As Long, ByVal lpFileName As String) As Long
'Private Declare Function CopyRect Lib "user32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Public Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As Rect, ByVal lpDescription As String) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
'Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DPtoLP Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
'Private Declare Function EnumEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEmf As Long, ByVal lpEnhMetaFunc As Long, lpData As Any, lpRect As Rect) As Long
'Private Declare Function EnumMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMetafile As Long, ByVal lpMFEnumProc As Long, ByVal lParam As Long) As Long
'Private Declare Function EqualRect Lib "user32" (lpRect1 As Rect, lpRect2 As Rect) As Long
'Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, ByVal FillValue As Byte)

'Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hEmf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
'Private Declare Function GetEnhMetaFile Lib "gdi32" Alias "GetEnhMetaFileA" (ByVal lpszMetaFile As String) As Long
'Private Declare Function GetEnhMetaFileDescription Lib "gdi32" Alias "GetEnhMetaFileDescriptionA" (ByVal hEmf As Long, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long
'Private Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hEmf As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
'Private Declare Function GetEnhMetaFilePaletteEntries Lib "gdi32" (ByVal hEmf As Long, ByVal cEntries As Long, lppe As PALETTEENTRY) As Long
Public Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hEmf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As Long) As Long
'Private Declare Function GetMetaFile Lib "gdi32" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
'Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
'Private Declare Function GetMetaRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetViewportExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As Size) As Long
Public Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
'Private Declare Function GetWindowExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As Size) As Long
'Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function GetWindowsDirectoryB Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub agCopyData Lib "apigid32.dll" (Source As Any, Dest As Any, ByVal nCount&)

'Private Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hfile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
'Private Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hfile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long

'Private Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function IntersectRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long
'Private Declare Function IsRectEmpty Lib "user32" (lpRect As Rect) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
'Private Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hfile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
'Private Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hfile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
'Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hfile As Long) As Long
'Private Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hfile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
'Private Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function LPtoDP Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

'Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Long
'Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEmf As Long, lpRect As Rect) As Long
'Private Declare Function PlayEnhMetaFileRecord Lib "gdi32" (ByVal hDC As Long, lpHandletable As HANDLETABLE, lpEnhMetaRecord As ENHMETARECORD, ByVal nHandles As Long) As Long
'Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long) As Long
'Private Declare Function PlayMetaFileRecord Lib "gdi32" (ByVal hDC As Long, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
'Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function ScaleViewportExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nXnum As Long, ByVal nXdenom As Long, ByVal nYnum As Long, ByVal nYdenom As Long, lpSize As Size) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Private Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Byte) As Long
'Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Public Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
'Private Declare Function SetMetaRgn Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
'Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function SetRectEmpty Lib "user32" (lpRect As Rect) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
'Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
'Private Declare Function SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtraA" (ByVal hDC As Long, ByVal nCharExtra As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Private Declare Function SetTextJustification Lib "gdi32" (ByVal hDC As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Public Declare Function SetViewportExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function SubtractRect Lib "user32" (lprcDst As Rect, lprcSrc1 As Rect, lprcSrc2 As Rect) As Long

'Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Private Declare Function UnionRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long

'Private Declare Function WaitForDoubleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
'Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
'
'Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Dest As Any, ByVal numBytes As Long)


'GradientFill constants, structures, declare

Public Const GRADIENT_FILL_RECT_H  As Long = &H0
'Public Const GRADIENT_FILL_RECT_V  As Long = &H1
'Public Const GRADIENT_FILL_TRIANGLE As Long = &H2
      
Public Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Public Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type
   
Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long


'winspool APIs

'Public Const CCHDEVICENAME = 32
'Public Const CCHFORMNAME = 32

'Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8


'Private Type ACL
'    AclRevision As Byte
'    Sbz1 As Byte
'    AclSize As Integer
'    AceCount As Integer
'    Sbz2 As Integer
'End Type

'Private Type SECURITY_DESCRIPTOR
'    Revision As Byte
'    Sbz1 As Byte
'    Control As Long
'    Owner As Long
'    Group As Long
'    Sacl As ACL
'    Dacl As ACL
'End Type

Public Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As Long            'address to DEVMODE structure
        DesiredAccess As Long
End Type

'Private Type DOC_INFO_1
'        pDocName As String
'        pOutputFile As String
'        pDatatype As String
'End Type

'Private Type DOC_INFO_2
'        pDocName As String
'        pOutputFile As String
'        pDatatype As String
'        dwMode As Long
'        JobId As Long
'End Type

'Private Type DOCINFO
'        cbSize As Long
'        lpszDocName As String
'        lpszOutput As String
'End Type

'Private Type PRINTER_INFO_1
'        flags As Long
'        pDescription As String
'        pName As String
'        pComment As String
'End Type

'Private Type PRINTER_INFO_2
'        pServerName As String
'        pPrinterName As String
'        pShareName As String
'        pPortName As String
'        pDriverName As String
'        pComment As String
'        pLocation As String
'        pDevMode As DEVMODE
'        pSepFile As String
'        pPrintProcessor As String
'        pDatatype As String
'        pParameters As String
'        pSecurityDescriptor As SECURITY_DESCRIPTOR
'        Attributes As Long
'        Priority As Long
'        DefaultPriority As Long
'        StartTime As Long
'        UntilTime As Long
'        Status As Long
'        cJobs As Long
'        AveragePPM As Long
'End Type

'Private Type PRINTER_INFO_3
'    pSecurityDescriptor As SECURITY_DESCRIPTOR
'End Type

'Private Type PRINTER_INFO_4
'    pPrinterName As String
'    pServerName As String
'    Attributes As Long
'End Type

'Private Type PRINTER_INFO_5
'    pPrinterName As String
'    pPortName As String
'    Attributes As Long
'    DeviceNotSelectedTimeout As Long
'    TransmissionRetryTimeout As Long
'End Type


Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE, ByVal fMode As Long) As Long
Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
'Private Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, pDevModeInput As DEVMODE) As Long
'Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long


' MonroeAdditions
'Public Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Public Declare Function PaintRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'Public Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As Rect) As Long



