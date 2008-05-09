Attribute VB_Name = "Module_Declaration"
Option Explicit

Public mTv As TreeView
Public mTVItems As New Collection

Public Const GIID_FIRST = 1
Public Const GIID_LAST = -1

'# MESSAGES
Public Const WM_USER = &H400
Public Const WM_DESTROY = &H2
Public Const WM_LBUTTONUP = &H202

Public Const NM_FIRST = -0&   ' (0U-  0U)       ' // generic to all controls
Public Const NM_LCLICK = (NM_FIRST - 2)
Public Const NM_DBLCLK = (NM_FIRST - 3)
Public Const NM_RETURN = (NM_FIRST - 4)
Public Const NM_RCLICK = (NM_FIRST - 5)

Public Const WM_NOTIFY = &H4E
' http://msdn.microsoft.com/library/devprods/vs6/visualc/vccore/_core_activex_controls.3a_.subclassing_a_windows_control.htm
Public Const OCM__BASE = (WM_USER + &H1C00)
Public Const OCM_NOTIFY = (OCM__BASE + WM_NOTIFY)

Public Const TVN_FIRST = -400&   ' (0U-400U)
Public Const TVN_SELCHANGED = (TVN_FIRST - 2)         ' lParam = NMTREEVIEW
Public Const TVN_ITEMEXPANDING = (TVN_FIRST - 5)    ' lParam = NMTREEVIEW
Public Const TVN_DELETEITEM = (TVN_FIRST - 9)           ' lParam = NMTREEVIEW

'# Structure renvoyee lors du message WM_NOTIFY
Public Type NMHDR
  hwndFrom As Long   '# Handle du controle a l'origine du message
  idFrom As Long        '# N° Identifiant du controle qui a envoye le message
  code  As Long          '# Code de notification
End Type

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Const TV_FIRST = &H1100
Public Const TVM_GETITEMRECT As Long = (TV_FIRST + 4)
Public Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Public Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Public Const TVM_GETITEM = (TV_FIRST + 12)
Public Const TVM_SETITEM = (TV_FIRST + 13)
Public Const TVM_HITTEST = (TV_FIRST + 17)
Public Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)

'# Permet de spécifier les checkboxes
Public Enum TVItemCheckStates
  TVICS_NONE = 0
  
  TVICS_BOLD_UNCHECKED = 1
  TVICS_BOLD_CHECKED = 2
  TVICS_UNCHECKED = 3
  TVICS_CHECKED = 4
  
  TVICS_BOLD_UNCROSSED = 5
  TVICS_BOLD_CROSSED = 6
  TVICS_UNCROSSED = 7
  TVICS_CROSSED = 8
  
  TVICS_BOLD_RADIO_UNCHECKED = 9
  TVICS_BOLD_RADIO_CHECKED = 10
  TVICS_RADIO_UNCHECKED = 11
  TVICS_RADIO_CHECKED = 12
End Enum

'# TVM_GET/SETIMAGELIST wParam
Public Const TVSIL_NORMAL = 0
Public Const TVSIL_STATE = 2

'# TVM_GETNEXTITEM wParam
Public Const TVGN_ROOT = &H0
Public Const TVGN_NEXT = &H1
Public Const TVGN_CHILD = &H4
Public Const TVGN_CARET = &H9

'# TVM_GET/SETITEM lParam
Public Type TVITEM
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

'# TVITEM mask
Public Const TVIF_TEXT = &H1
Public Const TVIF_IMAGE = &H2
Public Const TVIF_PARAM = &H4
Public Const TVIF_STATE = &H8
Public Const TVIF_HANDLE = &H10
Public Const TVIF_SELECTEDIMAGE = &H20
Public Const TVIF_CHILDREN = &H40
Public Const TVIF_INTEGRAL = &H80
Public Const TVIF_DI_SETITEM = &H1000

'# TVITEM state, stateMask
Public Const TVIS_SELECTED = &H2
Public Const TVIS_CUT = &H4
Public Const TVIS_DROPHILITED = &H8
Public Const TVIS_BOLD = &H10
Public Const TVIS_EXPANDED = &H20
Public Const TVIS_EXPANDEDONCE = &H40
Public Const TVIS_EXPANDPARTIAL = &H80
Public Const TVIS_OVERLAYMASK = &HF00
Public Const TVIS_STATEIMAGEMASK = &HF000

'# TVM_HITTEST lParam
Public Type TVHITTESTINFO
    pt As POINTAPI
    flags As TVHT_flags
    hItem As Long
End Type

'# Zone de clic lors du hottest
Public Enum TVHT_flags
    TVHT_NOWHERE = &H1
    TVHT_ONITEMICON = &H2
    TVHT_ONITEMLABEL = &H4
    TVHT_ONITEMINDENT = &H8
    TVHT_ONITEMBUTTON = &H10
    TVHT_ONITEMRIGHT = &H20
    TVHT_ONITEMSTATEICON = &H40
    TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
    TVHT_ONITEMLINE = (TVHT_ONITEM Or TVHT_ONITEMINDENT Or TVHT_ONITEMBUTTON Or TVHT_ONITEMRIGHT)
  
    TVHT_ABOVE = &H100
    TVHT_BELOW = &H200
    TVHT_TORIGHT = &H400
    TVHT_TOLEFT = &H800
End Enum

'# Utilise pour le tri sur CallBack
Public Type TVSORTCB
    hParent As Long
    lpfnCompare As Long
    lParam As Long
End Type

'# Strucutre pointée par lParam pour la plupart des messages de notification
Public Type NMTREEVIEW
    hdr As NMHDR
    action As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag As POINTAPI
End Type

'# IID des interfaces
Public Const IID_IShellFolder As String = "{000214E6-0000-0000-C000-000000000046}"
Public Const IID_IShellDetails As String = "{000214EC-0000-0000-C000-000000000046}"

'# TVITEM Index de l'Overlay
Public Const TVIO_SHARE = &H100
Public Const TVIO_SHORTCUT = &H200
Public Const TVIO_ARROW = &H300

Public Enum SpecialShellFolderIDs
  CSIDL_DESKTOP = &H0
  CSIDL_INTERNET = &H1
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
  CSIDL_ALTSTARTUP = &H1D
  CSIDL_COMMON_ALTSTARTUP = &H1E
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
'------------testé pour XP------------
  CSIDL_MYMUSICS = &HD
  CSIDL_WINDOWS = &H24
  CSIDL_SYSTEM32 = &H25
  CSIDL_PROGAMFILES = &H26
  CSIDL_MYPICTURES = &H27
  CSIDL_USERDIR = &H28
  CSIDL_COMMONFILES = &H2B
  CSIDL_SHAREDFOLDERS = &H2E
  CSIDL_ADMINTOOLS = &H2F
  CSIDL_NETWORKCONNECT = &H31
  CSIDL_SHAREDMUSICS = &H35
  CSIDL_SHAREDPICTURES = &H36
  CSIDL_RESOURCES = &H38
  CSIDL_CDBURNING = &H3B
  CSIDL_COMPUTERNETWORK = &H3D
End Enum

'# CodePage : utilise lors de la traduction / Unicode
Public Const CP_ACP = 0        ' ANSI code page
Public Const CP_OEMCP = 1   ' OEM code page

Public Const MB_PRECOMPOSED = &H1         '  use precomposed chars


Public Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long                '# PID de la racine
  pszDisplayName As String  '# Contient le nom du fichier selectionne
  lpszTitle As String              '# Legende
  ulFlags As Long                 '# Comportement de la fenetre
  lpfn As Long                     '# Adresse de la fonction de rappel
  lParam As Long                '# Info Supplémentaire qqui est passée a la fonction de rappel
  iImage As Long                '# variable ou l'on stocke l'image
End Type

Public Enum BF_Flags
    BIF_RETURNONLYFSDIRS = &H1              '# Renvoie un dossize
    BIF_DONTGOBELOWDOMAIN = &H2         '# Ne va pas plus loin sur le reseau
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_EDITBOX = &H10                              '# Ajoute une zone de saisie
    BIF_VALIDATE = &H20
    BIF_USENEWUI = &H40                            '# Nouvelle fenetre , elle permet de creer un nouveau repertoire, et peut etre agrandie
    BIF_BROWSEFORCOMPUTER = &H1000     '# Recherche d'ordinateur
    BIF_BROWSEFORPRINTER = &H2000         '# Recherche d'imprimant
    BIF_BROWSEINCLUDEFILES = &H4000      '# On inclue les fichiers
End Enum

'# Mesages vers la fonction de rappel
Public Enum BFFM_FromDlg
    BFFM_INITIALIZED = 1
    BFFM_SELCHANGED = 2
    BFFM_VALIDATEFAILEDA = 3
    BFFM_VALIDATEFAILEDW = 4
End Enum

'# Messages vers la fenetre du browser
Public Enum BFFM_ToDlg
    BFFM_SETSTATUSTEXTA = (WM_USER + 100)
    BFFM_ENABLEOK = (WM_USER + 101)
    BFFM_SETSELECTIONA = (WM_USER + 102)
    BFFM_SETSELECTIONW = (WM_USER + 103)
    BFFM_SETSTATUSTEXTW = (WM_USER + 104)
End Enum

Public Const MAX_PATH = 260
Public Const GWL_WNDPROC = (-4)

Public Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Public Enum SHGFI_Enum
    '#Icones
    SHGFI_LARGEICON = &H0
    SHGFI_SMALLICON = &H1
    SHGFI_OPENICON = &H2
    SHGFI_SHELLICONSIZE = &H4
  
    '# Comportement
    SHGFI_PIDL = &H8                            '# pszPath est le pidl , interdit l'acces au fichier par la fonction....
                                                        
    SHGFI_USEFILEATTRIBUTES = &H10   '# Suppose que le fichier existe
    SHGFI_ICON = &H100                        '# fills .hIcon, rtns BOOL, use DestroyIcon
    SHGFI_DISPLAYNAME = &H200           '# .szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
    SHGFI_TYPENAME = &H400                 '# Renvoie le type dans .szTypeName
    SHGFI_ATTRIBUTES = &H800              '# Renvoie les attributs
    SHGFI_ICONLOCATION = &H1000        '# place le nom, et l'icone dans fills .szDisplayName
    SHGFI_EXETYPE = &H2000                  '# renvoie deux caracteres ASCII indiquant le type d'Exe
    SHGFI_SYSICONINDEX = &H4000        '# .iIcon est l'indice systeme de l'icone
    SHGFI_LINKOVERLAY = &H8000           '# ajoutes le symbole de raccourcis sur .hicon
    SHGFI_SELECTED = &H10000              '# .hIcon est l'icone selectionnée
    SHGFI_ATTR_SPECIFIED = &H20000    '# n'extrait que les attributs spécifiés dans .dwAttributes
End Enum

