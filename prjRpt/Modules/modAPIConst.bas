Attribute VB_Name = "modAPIConst"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781F200A0"
Option Explicit
'**********************************************************************
'
'           Module Name: modAPIConst.bas
'
'           Purpose    : Declare constant used by WINDows API
'
'           Some Of the Constants are not Used in the Program
'           But you Can use it and purpose the These from the Documntation
'
'**********************************************************************

'for File IO API
Public Const INVALID_HANDLE_VALUE = -1
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_ALWAYS = 2
Public Const CREATE_NEW = 1
Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3
Public Const TRUNCATE_EXISTING = 5
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

'used by setfilepointer
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

'used by ShowOpen/ShowSave flags:
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

'Const for SetTextAlign
Public Const TA_BASELINE = 24 'The reference point will be on the baseline of the text.
Public Const TA_BOTTOM = 8 'The reference point will be on the bottom edge of the bounding rectangle of the text.
Public Const TA_CENTER = 6 'The reference point will be horizontally centered along the bounding rectangle of the text.
Public Const TA_LEFT = 0 'The reference point will be on the left edge of the bounding rectangle of the text.
Public Const TA_NOUPDATECP = 0 'Do not set the current point to the reference point.
Public Const TA_RIGHT = 2 'The reference point will be on the right edge of the bounding rectangle of the text.
Public Const TA_RTLREADING = 256 'Win 95/98 only:Display the text right-to-left (if the font is designed for right-to-left reading).
Public Const TA_TOP = 0 'The reference point will be on the top edge of the bounding rectangle of the text.
Public Const TA_UPDATECP = 1 'Set the current point to the reference point.

'PrintDlg flags:
Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80
Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

'used by GetDeviceCaps
Public Const TECHNOLOGY = 2 ' Device type returns DT_PLOTTER, DT_RASDISPLAY, DT_RASPRINTER, DT_RASCAMERA, DT_CHARSTREAM, DT_METAFILE, or DT_DISPFILE
Public Const DT_PLOTTER = 0 ' Vector plotter
Public Const DT_RASDISPLAY = 1 ' Raster display
Public Const DT_RASPRINTER = 2 ' Raster printer
Public Const DT_RASCAMERA = 3 ' Raster camera
Public Const DT_CHARSTREAM = 4 ' Character stream
Public Const VERTSIZE = 6 ' Width, in millimeters, of the physical screen.
Public Const HORZSIZE = 4 ' Height, in millimeters, of the physical screen.
Public Const HORZRES = 8  ' Width, in pixels, of the screen.
Public Const VERTRES = 10 ' Height, in raster lines, of the screen.
Public Const LOGPIXELSX = 88 ' (&H58) Number of pixels per logical inch along the screen width.
Public Const LOGPIXELSY = 90 ' (&H5A) Number of pixels per logical inch along the screen height.' For printing devices:
Public Const PHYSICALWIDTH = 110 ' (&H6E) The physical width, in device units.
Public Const PHYSICALHEIGHT = 111 ' (&H6F) The physical height, in device units.
Public Const PHYSICALOFFSETX = 112 ' (&H70) The physical printable area horizontal margin.
Public Const PHYSICALOFFSETY = 113 '(&H71) The physical printable area vertical margin.
Public Const SCALINGFACTORX = 114 ' (&H72)  The scaling factor along the horizontal axis.
Public Const SCALINGFACTORY = 115 ' (&H73  The scaling factor along the vertical axis.

'for DEVMODE type
Public Const DM_ORIENTATION = &H1
Public Const DM_PAPERSIZE = &H2
Public Const DM_PAPERLENGTH = &H4
Public Const DM_PAPERWIDTH = &H8
Public Const DM_SCALE = &H10
Public Const DM_COPIES = &H100
Public Const DM_DEFAULTSOURCE = &H200
Public Const DM_PRINTQUALITY = &H400
Public Const DM_COLOR = &H800
Public Const DM_DUPLEX = &H1000
Public Const DM_YRESOLUTION = &H2000
Public Const DM_TTOPTION = &H4000
Public Const DM_COLLATE = &H8000
Public Const DM_FORMNAME = &H10000
Public Const DM_LOGPIXELS = &H20000
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_ICMMETHOD = &H800000
Public Const DM_ICMINTENT = &H1000000
Public Const DM_MEDIATYPE = &H2000000
Public Const DM_DITHERTYPE = &H4000000
Public Const DM_PANNINGWIDTH = &H20000000
Public Const DM_PANNINGHEIGHT = &H40000000
Public Const DMORIENT_PORTRAIT = 1
Public Const DMORIENT_LANDSCAPE = 2
Public Const DMPAPER_LETTER = 1
Public Const DMPAPER_LEGAL = 5
Public Const DMPAPER_10X11 = 45
Public Const DMPAPER_10X14 = 16
Public Const DMPAPER_11X17 = 17
Public Const DMPAPER_15X11 = 46
Public Const DMPAPER_9X11 = 44
Public Const DMPAPER_A_PLUS = 57
Public Const DMPAPER_A2 = 66
Public Const DMPAPER_A3 = 8
Public Const DMPAPER_A3_EXTRA = 63
Public Const DMPAPER_A3_EXTRA_TRANSVERSE = 68
Public Const DMPAPER_A3_TRANSVERSE = 67
Public Const DMPAPER_A4 = 9
Public Const DMPAPER_A4_EXTRA = 53
Public Const DMPAPER_A4_PLUS = 60
Public Const DMPAPER_A4_TRANSVERSE = 55
Public Const DMPAPER_A4SMALL = 10
Public Const DMPAPER_A5 = 11
Public Const DMPAPER_A5_EXTRA = 64
Public Const DMPAPER_A5_TRANSVERSE = 61
Public Const DMPAPER_B_PLUS = 58
Public Const DMPAPER_B4 = 12
Public Const DMPAPER_B5 = 13
Public Const DMPAPER_B5_EXTRA = 65
Public Const DMPAPER_B5_TRANSVERSE = 62
Public Const DMPAPER_CSHEET = 24
Public Const DMPAPER_DSHEET = 25
Public Const DMPAPER_ENV_10 = 20
Public Const DMPAPER_ENV_11 = 21
Public Const DMPAPER_ENV_12 = 22
Public Const DMPAPER_ENV_14 = 23
Public Const DMPAPER_ENV_9 = 19
Public Const DMPAPER_ENV_B4 = 33
Public Const DMPAPER_ENV_B5 = 34
Public Const DMPAPER_ENV_B6 = 35
Public Const DMPAPER_ENV_C3 = 29
Public Const DMPAPER_ENV_C4 = 30
Public Const DMPAPER_ENV_C5 = 28
Public Const DMPAPER_ENV_C6 = 31
Public Const DMPAPER_ENV_C65 = 32
Public Const DMPAPER_ENV_DL = 27
Public Const DMPAPER_ENV_INVITE = 47
Public Const DMPAPER_ENV_ITALY = 36
Public Const DMPAPER_ENV_MONARCH = 37
Public Const DMPAPER_ENV_PERSONAL = 38
Public Const DMPAPER_ESHEET = 26
Public Const DMPAPER_EXECUTIVE = 7
Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41
Public Const DMPAPER_FANFOLD_STD_GERMAN = 40
Public Const DMPAPER_FANFOLD_US = 39
Public Const DMPAPER_FIRST = 1
Public Const DMPAPER_FOLIO = 14
Public Const DMPAPER_ISO_B4 = 42
Public Const DMPAPER_JAPANESE_POSTCARD = 43
Public Const DMPAPER_LAST = 41
Public Const DMPAPER_LEDGER = 4
Public Const DMPAPER_LEGAL_EXTRA = 51
Public Const DMPAPER_LETTER_EXTRA = 50

Public Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56
Public Const DMPAPER_LETTER_PLUS = 59
Public Const DMPAPER_LETTER_TRANSVERSE = 54
Public Const DMPAPER_LETTERSMALL = 2
Public Const DMPAPER_NOTE = 18
Public Const DMPAPER_QUARTO = 15

Public Const DMPAPER_STATEMENT = 6
Public Const DMPAPER_TABLOID = 3
Public Const DMPAPER_TABLOID_EXTRA = 52
Public Const DMPAPER_USER = 256
Public Const DMBIN_ONLYONE = 1
Public Const DMBIN_UPPER = 1
Public Const DMBIN_LOWER = 2
Public Const DMBIN_MIDDLE = 3
Public Const DMBIN_MANUAL = 4
Public Const DMBIN_ENVELOPE = 5
Public Const DMBIN_ENVMANUAL = 6
Public Const DMBIN_AUTO = 7
Public Const DMBIN_TRACTOR = 8
Public Const DMBIN_SMALLFMT = 9
Public Const DMBIN_LARGEFMT = 10
Public Const DMBIN_LARGECAPACITY = 11
Public Const DMBIN_CASSETTE = 14
Public Const DMBIN_FORMSOURCE = 15
Public Const DMRES_DRAFT = -1
Public Const DMRES_LOW = -2
Public Const DMRES_MEDIUM = -3
Public Const DMRES_HIGH = -4
Public Const DMCOLOR_MONOCHROME = 1
Public Const DMCOLOR_COLOR = 2
Public Const DMDUP_SIMPLEX = 1
Public Const DMDUP_VERTICAL = 2

Public Const DMDUP_HORIZONTAL = 3

Public Const DMTT_BITMAP = 1

Public Const DMTT_DOWNLOAD = 2

Public Const DMTT_SUBDEV = 4

Public Const DMCOLLATE_FALSE = 0

Public Const DMCOLLATE_TRUE = 1

Public Const DM_GRAYSCALE = 1

Public Const DM_INTERLACED = 2
Public Const DMICMMETHOD_NONE = 1
Public Const DMICMMETHOD_SYSTEM = 2
Public Const DMICMMETHOD_DRIVER = 3
Public Const DMICMMETHOD_DEVICE = 4
Public Const DMICM_SATURATE = 1
Public Const DMICM_CONTRAST = 2
Public Const DMICM_COLORMETRIC = 3
Public Const DMMEDIA_STANDARD = 1
Public Const DMMEDIA_GLOSSY = 2
Public Const DMMEDIA_TRANSPARECNY = 3

Public Const DMDITHER_NONE = 1

Public Const DMDITHER_COARSE = 2

Public Const DMDITHER_FINE = 3

Public Const DMDITHER_LINEART = 4

Public Const DMDITHER_GRAYSCALE = 5

'For Memory Related API

Public Const GHND = &H40
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_SHARE = &H2000
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = &H42
