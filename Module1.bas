Attribute VB_Name = "Module1"
Private Const mask = 2147483647
' Window Styles
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

'   Common Window Styles
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_CHILDWINDOW = (WS_CHILD)
' Extended Window Styles
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_TRANSPARENT = &H20&

' Dialog Styles
Private Const DS_ABSALIGN = &H1&
Private Const DS_SYSMODAL = &H2&
Private Const DS_LOCALEDIT = &H20          '  Edit items get Local storage.
Private Const DS_SETFONT = &H40            '  User specified font for Dlg controls
Private Const DS_MODALFRAME = &H80         '  Can be combined with WS_CAPTION
Private Const DS_NOIDLEMSG = &H100         '  WM_ENTERIDLE message will not be sent
Private Const DS_SETFOREGROUND = &H200     '  not in win3.1
Private Const DS_3DLOOK = &H4
Private Const DS_FIXEDSYS = &H8
Private Const DS_NOFAILCREATE = &H10
Private Const DS_CONTROL = &H400
Private Const DS_CENTER = &H800
Private Const DS_CENTERMOUSE = &H1000
Private Const DS_CONTEXTHELP = &H2000

Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87

' Predefined Resource Types
Private Const RT_CURSOR = 1
Private Const RT_BITMAP = 2
Private Const RT_ICON = 3
Private Const RT_MENU = 4
Private Const RT_DIALOG = 5
Private Const RT_STRING = 6
Private Const RT_FONTDIR = 7
Private Const RT_FONT = 8
Private Const RT_ACCELERATOR = 9
Private Const RT_RCDATA = 10
Private Const RT_MESSAGETABLE = 11
Private Const DIFFERENCE = 11
Private Const RT_GROUP_CURSOR = RT_CURSOR + DIFFERENCE
Private Const RT_GROUP_ICON = RT_ICON + DIFFERENCE
Private Const RT_VERSION = 16
Private Const RT_DLGINCLUDE = 17
Private Const RT_PLUGPLAY = 19
Private Const RT_VXD = 20
Private Const RT_ANICURSOR = 21
Private Const RT_ANIICON = 22
Private Const RT_HTML = 23

Private Const FVIRTKEY = &H1
Private Const FNOINVERT = &H2
Private Const FSHIFT = &H4
Private Const FCONTROL = &H8
Private Const FALT = &H10

Private Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Private Type IMAGEFILEHEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Private Type IMAGEDATADIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Private Type IMAGEOPTIONALHEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved1 As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(1 To 16) As IMAGEDATADIRECTORY
End Type
       
Private Type IMAGESECTIONHEADER
    NameSec As String * 8
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Private Type IMAGERESOURCEDIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    NumberOfNamedEntries As Integer
    NumberOfIdEntries As Integer
End Type

Private Type IMAGERESOURCEDIRECTORYENTRY
    Name As Long
    OffsetToData As Long
End Type

Private Type IMAGERESOURCEDATAENTRY
    OffsetToData As Long
    Size As Long
    CodePage As Long
    Reserved As Long
End Type

Private Type IMAGERESOURCEDIRSTRINGU
    Length As Integer
    NameString(64) As Byte
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type IMAGEIMPORTMODULEDIRECTORY
    dwRVAFunctionNameList As Long
    dwUseless1 As Long
    dwUseless2 As Long
    dwRVAModuleName As Long
    dwRVAFunctionAddressList As Long
End Type

Private Type GRPICONDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wPlanes As Integer
    wBitCount As Integer
    dwBytesInRes As Long
    nID As Integer
End Type

Private Type ICONDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wPlanes As Integer
    wBitCount As Integer
    dwBytesInRes As Long
    dwImageOffset As Long
End Type

Private Type GRPCURSORDIRENTRY
    wWidth As Integer
    wHeight As Integer
    wPlanes As Integer
    wBitCount As Integer
    lBytesInRes As Long
    wNameOrdinal As Integer
End Type

Private Type CURSORDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wHotspotX As Integer
    wHotspotY As Integer
    dwBytesInRes As Long
    dwImageOffset As Long
End Type

' version
Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Private Type VS_VERSIONINFO
    wLength As Integer
    wValueLength As Integer
    wType As Integer
    szKey(1 To 30) As Byte
    Padding1 As Long ' guessing
    FixedFileInfo As VS_FIXEDFILEINFO
    'Padding2 As Byte 'guessing
    'Children As Byte 'guessing
End Type

Private Type StringFileInfo
    wLength As Integer
    wValueLength As Integer
    wType As Integer
    szKey(1 To 30) As Byte
End Type

Private Type StringTable
    wLength As Integer
    wValueLength As Integer
    wType As Integer
    szKey(1 To 16) As Byte
    padding As Integer
End Type

Private Type VarFileInfo
    wLength As Integer
    wValueLength As Integer
    wType As Integer
    szKey(1 To 22) As Byte
    padding As Long
End Type

Private Type Var
    wLength As Integer
    wValueLength As Integer
    wType As Integer
    szKey(1 To 22) As Byte
    padding As Long
    Value1 As Integer
    Value2 As Integer
End Type

Private Type AccelTableEntry
    fFlags As Integer
    wAscii As Integer
    wId As Integer
    padding As Integer
End Type

Private Type BYTEARRAY
    Data() As Byte
End Type

Private Type DialogBoxHeader
    lStyle As Long
    lExtendedStyle As Long
    NumberOfItems As Integer
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    'MenuName As Integer
    'ClassName As Integer
    'szCaption() As Byte
    'wPointSize As Integer
    'szFontName() As Byte
End Type

Private Type DLGITEMTEMPLATE
    style As Long
    dwExtendedStyle As Long
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    id As Integer
End Type

Private Type DLGTEMPLATEEX
    dlgVer As Integer
    Signature As Integer
    helpID As Long
    exStyle As Long
    style As Long
    cDlgItems As Integer
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    menu As Integer
    windowClass As Integer
    title() As Byte
    weight As Integer
    bItalic As Integer
    fontname() As Byte
End Type
'INACTIVE
'DEFAULT
'RIGHTJUSTIFY
Private Const MF_POPUP = &H10&
Private Const MF_END = &H80
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_HELP = &H4000&
Private Const MF_HILITE = &H80&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_INSERT = &H0&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_SYSMENU = &H2000&
Private Const MF_UNCHECKED = &H0&
Private Const MF_UNHILITE = &H0&
Private Const MF_USECHECKBITMAPS = &H200&
Private Const MF_MOUSESELECT = &H8000&
Private Const MF_BITMAP = &H4&


Private Type MenuHeader
    wVersion As Integer
    cbHeaderSize As Integer
End Type

Private Type NORMALMENUITEM
    resInfo As Integer
    'menuText() As Byte
End Type

Private Type POPUPMENUITEM
    resInfo As Integer
    PopId As Integer
    'menuText() As Byte
End Type

Private Type MENUEXTEMPLATEHEADER
    wVersion As Integer
    wOffset As Integer
    dwHelpId As Long
End Type

Private Type MENUEXTEMPLATEITEM
    dwType As Long
    dwState As Long
    uId As Integer
    bResInfo As Integer
    szText() As Byte
    'dwHelpId As Long
End Type


Private Type MESSAGERESOURCEDATA
    NumberOfBlocks As Long
End Type

Private Type MESSAGERESOURCEBLOCK
    LowId As Long
    HighId As Long
    OffsetToEntries As Long
End Type

Private Type MESSAGERESOURCEENTRY
    Length As Integer
    flags As Integer
    Text(64) As Byte
End Type

Private Doshead As IMAGEDOSHEADER
Private Filehead As IMAGEFILEHEADER
Private ImgOphead As IMAGEOPTIONALHEADER
Private SectionHead As IMAGESECTIONHEADER

Private RootResDir As IMAGERESOURCEDIRECTORY
Private TypeIrde() As IMAGERESOURCEDIRECTORYENTRY

Private BmpResDir1 As IMAGERESOURCEDIRECTORY
Private BmpResDir2() As IMAGERESOURCEDIRECTORY
Private BmpResIrde1() As IMAGERESOURCEDIRECTORYENTRY
Private BmpResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Private BmpIRDAE() As IMAGERESOURCEDATAENTRY

Private IcoResDir1 As IMAGERESOURCEDIRECTORY
Private IcoResDir2() As IMAGERESOURCEDIRECTORY
Private IcoResIrde1() As IMAGERESOURCEDIRECTORYENTRY
Private IcoResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Private IcoIRDAE() As IMAGERESOURCEDATAENTRY

Private GrpIcoResDir1 As IMAGERESOURCEDIRECTORY
Private GrpIcoResDir2() As IMAGERESOURCEDIRECTORY
Private GrpIcoResIrde1() As IMAGERESOURCEDIRECTORYENTRY
Private GrpIcoResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Private GrpIcoIRDAE() As IMAGERESOURCEDATAENTRY

Private CurResDir1 As IMAGERESOURCEDIRECTORY
Private CurResDir2() As IMAGERESOURCEDIRECTORY
Private CurResIrde1() As IMAGERESOURCEDIRECTORYENTRY
Private CurResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Private CurIRDAE() As IMAGERESOURCEDATAENTRY

Private GrpCurResDir1 As IMAGERESOURCEDIRECTORY
Private GrpCurResDir2() As IMAGERESOURCEDIRECTORY
Private GrpCurResIrde1() As IMAGERESOURCEDIRECTORYENTRY
Private GrpCurResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Private GrpCurIRDAE() As IMAGERESOURCEDATAENTRY

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const OFN_HIDEREADONLY = &H4
Private Type OPENFILENAME
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


Dim BinName As String

Private Sub GetPEinfo()

Dim FileHandle As Long
Dim lnga As Long
Dim NewFileName As String
Dim FileNamo As String

Dim ResDir1 As IMAGERESOURCEDIRECTORY
Dim ResDir2() As IMAGERESOURCEDIRECTORY
Dim ResIrde1() As IMAGERESOURCEDIRECTORYENTRY

Dim TmpIrde1 As IMAGERESOURCEDIRECTORYENTRY
Dim ResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Dim IRDAE() As IMAGERESOURCEDATAENTRY

Dim RSRCsechead As Long

Dim IrdeOff As Long
Dim IrdeOffset As Long
Dim ResType As Long

FileHandle = FreeFile



Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
OpenFile.lStructSize = Len(OpenFile)
OpenFile.hwndOwner = 0
OpenFile.hInstance = App.hInstance
sFilter = "PE Header Files *.exe;*.dll;*.cpl;*.ocx;*.vxd;*.scr" & Chr(0) & "*.exe;*.dll;*.cpl;*.ocx;*.vxd;*.scr" & Chr(0)
OpenFile.lpstrFilter = sFilter
OpenFile.nFilterIndex = 1
OpenFile.lpstrFile = String(257, " ")
OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
OpenFile.lpstrFileTitle = OpenFile.lpstrFile
OpenFile.nMaxFileTitle = OpenFile.nMaxFile
OpenFile.lpstrInitialDir = "C:\windows\"
OpenFile.lpstrTitle = "Portable Executable UnNamed"
OpenFile.flags = OFN_HIDEREADONLY
lReturn = GetOpenFileName(OpenFile)
If lReturn = 0 Then
    End
Else
    FileNamo = Trim(OpenFile.lpstrFile)
End If
BinName = Mid(FileNamo, InStrRev(FileNamo, "\") + 1, Len(FileNamo) - InStrRev(FileNamo, "\") - 1)
Open FileNamo For Binary As FileHandle
Get FileHandle, 1, Doshead
Get FileHandle, Doshead.e_lfanew, lnga
  If Hex(lnga) = "455000" Then
Else
       MsgBox "Not a valid Portable Executable File" & Hex(lnga), vbCritical
       Exit Sub
End If
Get FileHandle, Doshead.e_lfanew + Len(lnga) + 1, Filehead
Get FileHandle, Doshead.e_lfanew + Len(lnga) + Len(Filehead) + 1, ImgOphead
Get FileHandle, Doshead.e_lfanew + Len(lnga) + Len(Filehead) + Len(ImgOphead) + 1, SectionHead
       Do While i < Filehead.NumberOfSections
                If Mid(SectionHead.NameSec, 1, 5) = ".rsrc" Then
                    Exit Do
                End If
                RSRCsechead = RSRCsechead + 40
                Get FileHandle, Doshead.e_lfanew + Len(lnga) + Len(Filehead) + Len(ImgOphead) + 1 + RSRCsechead, SectionHead
                i = i + 1
        Loop
Get FileHandle, SectionHead.PointerToRawData + 1, RootResDir
ReDim TypeIrde(1 To RootResDir.NumberOfIdEntries + RootResDir.NumberOfNamedEntries)
For ResType = 1 To RootResDir.NumberOfIdEntries + RootResDir.NumberOfNamedEntries
IrdeOff = 0
    Get FileHandle, SectionHead.PointerToRawData + 1 + Len(RootResDir) + IrdeOffset, TypeIrde(ResType)
    Get FileHandle, (mask + TypeIrde(ResType).OffsetToData + 2) + SectionHead.PointerToRawData, ResDir1
    ReDim ResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    ReDim ResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    ReDim ResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    ReDim IRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    Dim counter As Long
    For i = 1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries
        Get FileHandle, (mask + TypeIrde(ResType).OffsetToData + 2) + SectionHead.PointerToRawData + Len(ResDir1) + IrdeOff, ResIrde1(i)
        Get FileHandle, mask + ResIrde1(i).OffsetToData + 2 + SectionHead.PointerToRawData, ResDir2(i)
        If (ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries) < (ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries) Then
            ReDim ResIrde2(1 To ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries)
            ReDim IRDAE(1 To ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries)
            counter = 1
        Else
            counter = i
        End If
        If (ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries) > 1 Then
            counter = 1
        End If
        Dim newLoop As Long
        newLoop = 0
        Dim NewOffset As Long
        NewOffset = 0
        For newLoop = 0 To ((ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries) - 1)
            Get FileHandle, mask + ResIrde1(i).OffsetToData + 2 + SectionHead.PointerToRawData + Len(ResDir1) + NewOffset, ResIrde2(newLoop + counter)
            Get FileHandle, ResIrde2(newLoop + counter).OffsetToData + SectionHead.PointerToRawData + 1, IRDAE(newLoop + counter)
            Select Case TypeIrde(ResType).Name
                Case Is < 0
                    Dim ResTypeName As String
                    ResTypeName = GetResourceName(TypeIrde(ResType).Name, FileHandle)
                    Select Case ResTypeName
                Case "AVI", "ANIMATION"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".avi"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "WAVE", "SOUND"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".wav"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "RCDATA", "TCOMPRESS"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".rcdata"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "TEXT", "MSCONTROLS"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".rtf"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "REGINST", "REGISTRY"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".regtxt"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "TYPELIB"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".typelib"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "BINARY"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".bin"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case "EXEFILE", "EXE"
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    NewFileName = NewFileName + ".exe"
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
                Case Else
                    NewFileName = (GetResourceName(ResIrde1(i).Name, FileHandle))
                    NewFileName = ResTypeName + NewFileName
                    DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
            End Select
            Case RT_BITMAP
                ReDim BmpResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim BmpResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim BmpResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim BmpIRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                BmpResDir1 = ResDir1
                BmpResIrde1 = ResIrde1
                BmpResDir2 = ResDir2
                BmpResIrde2 = ResIrde2
                BmpIRDAE = IRDAE
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                SaveResBitmap IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, FileHandle, NewFileName, IRDAE(newLoop + counter).Size
            Case RT_GROUP_ICON
                ReDim GrpIcoResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim GrpIcoResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim GrpIcoResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim GrpIcoIRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                GrpIcoResDir1 = ResDir1
                GrpIcoResIrde1 = ResIrde1
                GrpIcoResDir2 = ResDir2
                GrpIcoResIrde2 = ResIrde2
                GrpIcoIRDAE = IRDAE
                If IcoResDir1.NumberOfIdEntries + IcoResDir1.NumberOfNamedEntries > 0 Then
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                End If
                SaveResIcon IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, FileHandle, NewFileName, IRDAE(newLoop + counter).Size
            Case RT_GROUP_CURSOR
                ReDim GrpCurResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim GrpCurResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim GrpCurResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim GrpCurIRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                GrpCurResDir1 = ResDir1
                GrpCurResIrde1 = ResIrde1
                GrpCurResDir2 = ResDir2
                GrpCurResIrde2 = ResIrde2
                GrpCurIRDAE = IRDAE
                If CurResDir1.NumberOfIdEntries + CurResDir1.NumberOfNamedEntries > 0 Then
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                End If
                    SaveResCursor IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, FileHandle, NewFileName, IRDAE(newLoop + counter).Size
              Case RT_ICON
                ReDim IcoResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim IcoResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim IcoResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim IcoIRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                IcoResDir1 = ResDir1
                IcoResIrde1 = ResIrde1
                IcoResDir2 = ResDir2
                IcoResIrde2 = ResIrde2
                IcoIRDAE = IRDAE
                If GrpIcoResDir1.NumberOfIdEntries > 0 Then
                    SaveResIcon IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, FileHandle, CStr(ResIrde1(i).Name), IRDAE(newLoop + counter).Size
                End If
              Case RT_CURSOR
                ReDim CurResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim CurResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim CurResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                ReDim CurIRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
                CurResDir1 = ResDir1
                CurResIrde1 = ResIrde1
                CurResDir2 = ResDir2
                CurResIrde2 = ResIrde2
                CurIRDAE = IRDAE
                If GrpCurResDir1.NumberOfIdEntries > 0 Then
                    SaveResCursor IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, FileHandle, CStr(ResIrde1(i).Name), IRDAE(newLoop + counter).Size
                End If
              Case RT_VERSION
                      DumpVersionInfo IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, FileHandle
              Case RT_STRING
                    DumpStringTable IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, ResIrde1(i).Name, IRDAE(newLoop + counter).Size, FileHandle
              Case RT_MENU
                NewFileName = (GetResourceName(ResIrde1(i).Name, FileHandle))
                DumpMenuTable2 IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
              Case RT_ACCELERATOR
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                DumpAccelTable IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
              Case RT_DIALOG
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
               ' DumpDialogTable IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
              Case RT_MESSAGETABLE
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                DumpMessageTable IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, IRDAE(newLoop + counter).Size, FileHandle
              Case RT_VXD
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                NewFileName = NewFileName + ".vxd"
                DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
              Case RT_HTML
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                If InStr(1, NewFileName, ".") = 0 Then
                    NewFileName = NewFileName + ".html"
                End If
                DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
              Case Else
                NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                DumpAnyResource IRDAE(newLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, NewFileName, IRDAE(newLoop + counter).Size, FileHandle
            End Select
            NewOffset = NewOffset + Len(ResIrde2(newLoop + counter))
        Next newLoop
    IrdeOff = IrdeOff + Len(ResIrde2(i))
    Next i
IrdeOffset = IrdeOffset + Len(TypeIrde(ResType))
Next ResType
Close FileHandle
Exit Sub
Err:
MsgBox "Errors occured while opeing the file : " & FileNamo, vbCritical
End Sub

Private Function SaveResBitmap(Offset As Long, OpenResFile As Long, ResName As String, ResSize As Long) As Boolean
Dim BMPHDR As BITMAPINFOHEADER
Dim BmpFileHdr As BITMAPFILEHEADER
Dim SaveFileName As String
Dim strBmpData() As Byte
 Get OpenResFile, Offset, BMPHDR
        If BMPHDR.biBitCount <= 8 Then
                If BMPHDR.biClrUsed <> 0 Then
                TableSize = BMPHDR.biClrUsed
        Else
                TableSize = 2 ^ BMPHDR.biBitCount
                End If
        End If
        With BmpFileHdr
            .bfType = &H4D42
            .bfSize = ResSize + 14
            .bfOffBits = Len(BmpFileHdr) + Len(BMPHDR) + TableSize * 4
        End With
        filenum = FreeFile
        ReDim strBmpData(1 To ResSize)
        Get OpenResFile, Offset, strBmpData
        SaveFileName = "c:\6110\" & BinName & "." & ResName & ".bmp"
        Open SaveFileName For Binary Lock Write As #filenum
            Put #filenum, , BmpFileHdr
            Put #filenum, , strBmpData
        Close #filenum
End Function

Private Function SaveResIcon(Offset As Long, OpenResFile As Long, ResName As String, ResSize As Long) As Boolean
Dim GroupData() As Byte
Dim IconCount As Integer
Dim SaveFileName As String
Dim GrpIconEntries() As GRPICONDIRENTRY
Dim CurEntry As ICONDIRENTRY
Dim IcoData() As Byte
Dim IcoOffset As Long
ReDim GroupData(0 To 5)
    filenum = FreeFile
    Get OpenResFile, Offset, GroupData
    SaveFileName = "c:\6110\" & BinName & "." & ResName & ".ico"
    Open SaveFileName For Binary Lock Write As #filenum
    IconCount = GroupData(4)
    Put #filenum, , GroupData
    ReDim GrpIconEntries(0 To IconCount - 1)
For i = 0 To IconCount - 1
    Get OpenResFile, Offset + 6 + i * 14, GrpIconEntries(i)
Next i
    IcoOffset = 6 + Len(CurEntry) * IconCount
    For N = 0 To IconCount - 1
        With GrpIconEntries(N)
                CurEntry.bWidth = .bWidth
                CurEntry.bHeight = .bHeight
                CurEntry.bColorCount = .bColorCount
                CurEntry.wPlanes = .wPlanes
                CurEntry.wBitCount = .wBitCount
                CurEntry.dwBytesInRes = .dwBytesInRes
                CurEntry.dwImageOffset = IcoOffset
                IcoOffset = IcoOffset + .dwBytesInRes
        End With
        Put #filenum, , CurEntry
    Next N
    N = 0
    i = 1
    For N = 0 To IconCount - 1
        For i = 0 To IcoResDir1.NumberOfIdEntries - 1
            If IcoResIrde1(i + 1).Name = GrpIconEntries(N).nID Then
                ReDim IcoData(1 To IcoIRDAE(i + 1).Size)
                Get OpenResFile, IcoIRDAE(i + 1).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, IcoData
                Put #filenum, , IcoData
                GoTo NextN
            End If
        Next i
NextN:
    Next N
    Close filenum
End Function

Private Function SaveResCursor(Offset As Long, OpenResFile As Long, ResName As String, ResSize As Long) As Boolean
Dim GroupData() As Byte
Dim CursorCount As Integer
Dim SaveFileName As String
Dim GrpCursorEntries() As GRPCURSORDIRENTRY
Dim GrpCursorEntriesOg() As GRPCURSORDIRENTRY
Dim CurEntry As CURSORDIRENTRY
Dim FinalCurData() As BYTEARRAY
Dim CurData() As String
Dim CurOffset As Long
ReDim GroupData(0 To 5)
Dim CurHsptX As Integer, CurHsptY As Integer
Dim N As Long
    filenum = FreeFile
    Get OpenResFile, Offset, GroupData
    SaveFileName = "c:\6110\" & BinName & "." & ResName & ".cur"
    Open SaveFileName For Binary Lock Write As #filenum
    CursorCount = GroupData(4)
    Put #filenum, , GroupData
    ReDim GrpCursorEntries(0 To CursorCount - 1)
    ReDim GrpCursorEntriesOg(0 To CursorCount - 1)
    ReDim CurData(0 To CursorCount - 1)
For i = 0 To CursorCount - 1
    Get OpenResFile, Offset + 6 + i * 14, GrpCursorEntries(i)
Next i
i = 0
    CurOffset = 6 + Len(CurEntry) * CursorCount
    For N = 0 To CursorCount - 1
        With GrpCursorEntries(N)
            For i = 0 To CurResDir1.NumberOfIdEntries - 1
                If CurResIrde1(i + 1).Name = GrpCursorEntries(N).wNameOrdinal Then
                    Get OpenResFile, CurIRDAE(i + 1).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, CurHsptX
                    Get OpenResFile, CurIRDAE(i + 1).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1 + 2, CurHsptY
                    CurData(N) = String(CurIRDAE(i + 1).Size - 4, " ")
                    Get OpenResFile, CurIRDAE(i + 1).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1 + 4, CurData(N)
                    GoTo NextCI
                End If
            Next i
NextCI:
            CurEntry.bWidth = .wWidth
            CurEntry.bHeight = .wHeight / 2
            CurEntry.wHotspotX = CurHsptX
            CurEntry.wHotspotY = CurHsptY
            CurEntry.dwBytesInRes = .lBytesInRes - 4
            CurEntry.dwImageOffset = CurOffset
            CurOffset = CurOffset + .lBytesInRes - 4
        End With
        Put #filenum, , CurEntry
    Next N
    i = 1
    For j = 0 To CursorCount - 1
       Put #filenum, , CurData(j)
    Next j
    Close filenum
End Function

Private Function GetResourceName(Offset As Long, OpenResFile As Long) As Byte()
  Dim tmpstr As IMAGERESOURCEDIRSTRINGU
  Dim stra() As Byte
If Offset < 0 Then
  Offset = mask + Offset + SectionHead.PointerToRawData + 2
  Get OpenResFile, Offset, tmpstr
  ReDim stra(tmpstr.Length * 2)
  Get OpenResFile, Offset + 2, stra
  GetResourceName = Mid(stra, 1, tmpstr.Length)
Else
    GetResourceName = CStr(Offset)
End If
End Function

Private Function DumpVersionInfo(Offset As Long, OpenResFile As Long)
Dim FileVerInfo As VS_VERSIONINFO
Dim StrFileInfo As StringFileInfo
Dim StrTable As StringTable
Dim VarInfo As VarFileInfo
Dim VarInfoVar As Var
Dim IntStrLoop As Integer
Dim IntTotLoop As Integer
Dim IntData As Integer
Dim strLen As Integer
Dim strInfo(1 To 28) As Byte
Dim ByData() As Byte
Dim StrData As String
Dim tmpstr As String
Dim StrDataDump As String
Dim freenum As Integer

freenum = FreeFile
Open "c:\6110\" & BinName & "." & "version.txt" For Binary As freenum

Get OpenResFile, Offset, FileVerInfo
IntTotLoop = Len(FileVerInfo)
Put freenum, LOF(freenum) + 1, CStr(FileVerInfo.szKey) & vbNewLine
Put freenum, , "   " & CStr(Hex(FileVerInfo.FixedFileInfo.dwSignature)) & vbNewLine
' to input later
Do While IntTotLoop < FileVerInfo.wLength
    Get OpenResFile, Offset + IntTotLoop + 6, strInfo
        If CStr(strInfo()) = "StringFileInfo" Then
            Get OpenResFile, Offset + IntTotLoop, StrFileInfo
            Put freenum, , CStr(StrFileInfo.szKey) & vbNewLine
            Get OpenResFile, Offset + IntTotLoop + Len(StrFileInfo), IntData
                If IntData = 0 Then
                    IntStrLoop = IntStrLoop + 2
                End If
            Get OpenResFile, Offset + IntTotLoop + Len(StrFileInfo) + IntStrLoop, StrTable
            Put freenum, , "   " & CStr(StrTable.szKey) & vbNewLine
            IntStrLoop = Len(StrTable)
            Do While IntStrLoop < StrTable.wLength
                Get OpenResFile, Offset + IntTotLoop + Len(StrFileInfo) + IntStrLoop, IntData
                If IntData = 0 Or IntData > StrTable.wLength Or IntData < 0 Then
                    IntStrLoop = IntStrLoop + 1
                Else
                    StrData = String(IntData - 6, " ")
                    Get OpenResFile, Offset + IntTotLoop + Len(StrFileInfo) + IntStrLoop + 6, StrData
                    Get OpenResFile, Offset + IntTotLoop + Len(StrFileInfo) + IntStrLoop + 2, strLen
                    'StrDataDump = StrConv(StrData, vbFromUnicode)
                    'MsgBox Len(strdata2)
                   ' StrDataDump = Left(StrDataDump, InStr(1, StrDataDump, vbNullChar, vbTextCompare) - 1)
                    StrData = Trim(StrConv(StrData, vbFromUnicode))
                    
                        Put freenum, , "      " & StrData & vbNewLine
                
                    IntStrLoop = IntStrLoop + IntData
                End If
                If IntStrLoop = StrTable.wLength Then
                    IntTotLoop = IntTotLoop + StrFileInfo.wLength
                ElseIf IntStrLoop > StrTable.wLength Then
                    MsgBox "Error in Dumping Version Info", vbCritical
                    Exit Function
                End If
                    Loop
                End If
        If Mid(CStr(strInfo()), 1, 11) = "VarFileInfo" Then
            Get OpenResFile, Offset + IntTotLoop, VarInfo
            Get OpenResFile, Offset + IntTotLoop + Len(VarInfo), VarInfoVar
            Put freenum, , CStr(VarInfo.szKey) & vbNewLine
            Put freenum, , "    " & CStr(VarInfoVar.szKey) & vbNewLine
            Put freenum, , "        " & "Language/Codepage :" & VarInfoVar.Value1 & "/" & VarInfoVar.Value2 & vbNewLine & vbNewLine
            IntTotLoop = IntTotLoop + VarInfo.wLength
        End If
Get OpenResFile, Offset + IntTotLoop, IntData
If IntData = 0 Then
    IntTotLoop = IntTotLoop + 2
End If
Loop
Close freenum
End Function
Private Function DumpStringTable(Offset As Long, StringTableName As Long, StringTableSize As Long, OpenResFile As Long)
Dim StringName As Long
Dim StringOffset As Integer
Dim IntCtr As Integer
Dim StringData As String
Dim freenum As Integer

freenum = FreeFile
Open "c:\6110\" & BinName & "." & "stringdata.txt" For Binary As freenum

StringName = StringTableName * 16 - 16
Put freenum, LOF(freenum) + 1, StringTableName & vbNewLine

For IntCtr = 1 To 16
    Get OpenResFile, Offset, StringOffset
    If StringOffset = 0 Then
        StringName = StringName + 1
        Offset = Offset + 2
    Else
        StringData = String(StringOffset * 2, " ")
        Get OpenResFile, Offset + 2, StringData
        StringData = StrConv(StringData, vbFromUnicode)
        Put freenum, , "   " & StringName & ",  " & StringData & vbNewLine
        Offset = Offset + StringOffset * 2 + 2
        StringName = StringName + 1
    End If
Next IntCtr
Close freenum
End Function
Private Function DumpAccelTable(Offset As Long, AccelTableName As String, AccelTableSize As Long, OpenResFile As Long)
Dim AccTable As AccelTableEntry
Dim ctr As Integer
Dim DumpString As String
Dim Offy As Long
Dim freenum As Integer

freenum = FreeFile
Open "c:\6110\" & BinName & "." & "accelerator.txt" For Binary As freenum
Put freenum, LOF(freenum) + 1, AccelTableName & vbNewLine
For ctr = 1 To AccelTableSize / 8
    Get OpenResFile, Offset + Offy, AccTable
    If ctr = (AccelTableSize / 8) Then
        AccTable.fFlags = AccTable.fFlags - 128
    End If
    Select Case AccTable.wAscii
        Case vbKeyLButton
            DumpString = "VK_LBUTTON"
        Case vbKeyRButton
            DumpString = "VK_RBUTTON"
        Case vbKeyCancel
            DumpString = "VK_CANCEL"
        Case vbKeyMButton
            DumpString = "VK_MBUTTON"
        Case vbKeyBack
            DumpString = "VK_BACK"
        Case vbKeyTab
            DumpString = "VK_TAB"
        Case vbKeyClear
            DumpString = "VK_CLEAR"
        Case vbKeyReturn
            DumpString = "VK_RETURN"
        Case vbKeyShift
            DumpString = "VK_SHIFT"
        Case vbKeyMenu
            DumpString = "VK_MENU"
        Case vbKeyPause
            DumpString = "VK_PAUSE"
        Case vbKeyCapital
            DumpString = "VK_CAPITAL"
        Case vbKeyEscape
            DumpString = "VK_ESCAPE"
        Case vbKeySpace
            DumpString = "VK_SPACE"
        Case vbkeyPRIOR
            DumpString = "VK_PRIOR"
        Case vbkeyNEXT
            DumpString = "VK_NEXT"
        Case vbKeyEnd
            DumpString = "VK_END"
        Case vbKeyHome
            DumpString = "VK_HOME"
        Case vbKeyLeft
            DumpString = "VK_LEFT"
        Case vbKeyUp
            DumpString = "VK_UP"
        Case vbKeyRight
            DumpString = "VK_RIGHT"
        Case vbKeyDown
            DumpString = "VK_DOWN"
        Case vbKeySelect
            DumpString = "VK_SELECT"
        Case vbKeyPrint
            DumpString = "VK_PRINT"
        Case vbKeyExecute
            DumpString = "VK_EXECUTE"
        Case vbKeySnapshot
            DumpString = "VK_SNAPSHOT"
        Case vbKeyInsert
            DumpString = "VK_INSERT"
        Case vbKeyDelete
            DumpString = "VK_DELETE"
        Case vbKeyHelp
            DumpString = "VK_HELP"
        Case vbKeyNumpad0
            DumpString = "VK_NUMPAD0"
        Case vbKeyNumpad1
            DumpString = "VK_NUMPAD1"
        Case vbKeyNumpad2
            DumpString = "VK_NUMPAD2"
        Case vbKeyNumpad3
            DumpString = "VK_NUMPAD3"
        Case vbKeyNumpad4
            DumpString = "VK_NUMPAD4"
        Case vbKeyNumpad5
            DumpString = "VK_NUMPAD5"
        Case vbKeyNumpad6
            DumpString = "VK_NUMPAD6"
        Case vbKeyNumpad7
            DumpString = "VK_NUMPAD7"
        Case vbKeyNumpad8
            DumpString = "VK_NUMPAD8"
        Case vbKeyNumpad9
            DumpString = "VK_NUMPAD9"
        Case vbKeyMultiply
            DumpString = "VK_MULTIPLY"
        Case vbKeyAdd
            DumpString = "VK_ADD"
        Case vbKeySeparator
            DumpString = "VK_SEPARATOR"
        Case vbKeySubtract
            DumpString = "VK_SUBTRACT"
        Case vbKeyDecimal
            DumpString = "VK_DECIMAL"
        Case vbKeyDivide
            DumpString = "VK_DIVIDE"
        Case vbKeyF1
            DumpString = "VK_F1"
        Case vbKeyF2
            DumpString = "VK_F2"
        Case vbKeyF3
            DumpString = "VK_F3"
        Case vbKeyF4
            DumpString = "VK_F4"
        Case vbKeyF5
            DumpString = "VK_F5"
        Case vbKeyF6
            DumpString = "VK_F6"
        Case vbKeyF7
            DumpString = "VK_F7"
        Case vbKeyF8
            DumpString = "VK_F8"
        Case vbKeyF9
            DumpString = "VK_F9"
        Case vbKeyF10
            DumpString = "VK_F10"
        Case vbKeyF11
            DumpString = "VK_F11"
        Case vbKeyF12
            DumpString = "VK_F12"
        Case vbKeyF13
            DumpString = "VK_F13"
        Case vbKeyF14
            DumpString = "VK_F14"
        Case vbKeyF15
            DumpString = "VK_F15"
        Case vbKeyF16
            DumpString = "VK_F16"
        Case VK_F17
            DumpString = "VK_F17"
        Case VK_F18
            DumpString = "VK_F18"
        Case VK_F19
            DumpString = "VK_F19"
        Case VK_F20
            DumpString = "VK_F20"
        Case VK_F21
            DumpString = "VK_F21"
        Case VK_F22
            DumpString = "VK_F22"
        Case VK_F23
            DumpString = "VK_F23"
        Case VK_F24
            DumpString = "VK_F24"
        Case vbKeyNumlock
            DumpString = "VK_NUMLOCK"
        Case vbkeySCROLL
            DumpString = "VK_SCROLL"
        Case Else
            DumpString = """" & Chr(AccTable.wAscii) & """"
    End Select
    If AccTable.wId < 1 Then
        DumpString = DumpString & ", " & CStr(HexToDecimal(Hex(AccTable.wId)))
    Else
        DumpString = DumpString & ", " & CStr(AccTable.wId)
    End If
    If AccTable.fFlags > 0 Then
        If (AccTable.fFlags And FVIRTKEY) = FVIRTKEY Then
            DumpString = DumpString + ", VIRTKEY"
        End If
        If (AccTable.fFlags And FNOINVERT) = FNOINVERT Then
            DumpString = DumpString + ", NOINVERT"
        End If
        If (AccTable.fFlags And FSHIFT) = FSHIFT Then
            DumpString = DumpString + ", SHIFT"
        End If
        If (AccTable.fFlags And FCONTROL) = FCONTROL Then
            DumpString = DumpString + ", CONTROL"
        End If
        If (AccTable.fFlags And FALT) = FALT Then
            DumpString = DumpString + ", ALT"
        End If
    Else
        DumpString = DumpString + ","
    End If
    Put freenum, LOF(freenum) + 1, "   " & DumpString & vbNewLine
    Offy = Offy + 8
Next ctr
Close freenum
End Function
Private Function HexToDecimal(RawString As String) As Long
RawString = "&H" & RawString & "&"
HexToDecimal = Val(RawString)
End Function
Private Function DumpDialogTable(Offset As Long, DialogName As String, DialogSize As Long, OpenResFile As Long)
Dim Signature As Integer
Dim StrTmp As String
Dim IntTmp As Integer
Dim DlgBxCaption As String
Dim DlgBxFontName As String
Dim DlgbxFontSize As Integer
Dim lngValDialog As Long
Get OpenResFile, Offset + 2, Signature
If Signature = -1 Then
    Dim DlgBxEx As DLGTEMPLATEEX
Else
    Dim DlgBx As DialogBoxHeader
    Get OpenResFile, Offset, DlgBx
    Get OpenResFile, Offset + Len(DlgBx), IntTmp
    If IntTmp <> 0 Then
        MsgBox "Please Check Menu NAME"
    End If
    Get OpenResFile, Offset + Len(DlgBx) + 2, IntTmp
    If IntTmp <> 0 Then
        MsgBox "Please Check Class NAME"
    End If
    StrTmp = String(255, " ")
    Get OpenResFile, Offset + Len(DlgBx) + 4, StrTmp
    IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar, vbTextCompare)
    DlgBxCaption = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp)
    Debug.Print DialogName
    Debug.Print DlgBxCaption
If (DlgBx.lStyle And DS_ABSALIGN) = DS_ABSALIGN Then
    strdlgstly = strdlgstly + " | DS_ABSALIGN"
    lngValDialog = lngValDialog + DS_ABSALIGN
End If
If (DlgBx.lStyle And DS_SYSMODAL) = DS_SYSMODAL Then
    strdlgstly = strdlgstly + " | DS_SYSMODAL"
    lngValDialog = lngValDialog + DS_SYSMODAL
End If
If (DlgBx.lStyle And DS_LOCALEDIT) = DS_LOCALEDIT Then
    strdlgstly = strdlgstly + " | DS_LOCALEDIT"
    lngValDialog = lngValDialog + DS_LOCALEDIT
End If
If (DlgBx.lStyle And DS_SETFONT) = DS_SETFONT Then
    strdlgstly = strdlgstly + " | DS_SETFONT"
    lngValDialog = lngValDialog + DS_SETFONT
End If
If (DlgBx.lStyle And DS_MODALFRAME) = DS_MODALFRAME Then
    strdlgstly = strdlgstly + " | DS_MODALFRAME"
    lngValDialog = lngValDialog + DS_MODALFRAME
End If
If (DlgBx.lStyle And DS_NOIDLEMSG) = DS_NOIDLEMSG Then
    strdlgstly = strdlgstly + " | DS_NOIDLEMSG"
    lngValDialog = lngValDialog + DS_NOIDLEMSG
End If
If (DlgBx.lStyle And DS_SETFOREGROUND) = DS_SETFOREGROUND Then
    strdlgstly = strdlgstly + " | DS_SETFOREGROUND"
    lngValDialog = lngValDialog + DS_SETFOREGROUND
End If
If (DlgBx.lStyle And DS_3DLOOK) = DS_3DLOOK Then
    strdlgstly = strdlgstly + " | DS_3DLOOK"
    lngValDialog = lngValDialog + DS_3DLOOK
End If
If (DlgBx.lStyle And DS_FIXEDSYS) = DS_FIXEDSYS Then
    strdlgstly = strdlgstly + " | DS_FIXEDSYS"
    lngValDialog = lngValDialog + DS_FIXEDSYS
End If
If (DlgBx.lStyle And DS_NOFAILCREATE) = DS_NOFAILCREATE Then
    strdlgstly = strdlgstly + " | DS_NOFAILCREATE"
    lngValDialog = lngValDialog + DS_NOFAILCREATE
End If
If (DlgBx.lStyle And DS_CONTROL) = DS_CONTROL Then
    strdlgstly = strdlgstly + " | DS_CONTROL"
    lngValDialog = lngValDialog + DS_CONTROL
End If
If (DlgBx.lStyle And DS_CENTER) = DS_CENTER Then
    strdlgstly = strdlgstly + " | DS_CENTER"
    lngValDialog = lngValDialog + DS_CENTER
End If
If (DlgBx.lStyle And DS_CENTERMOUSE) = DS_CENTERMOUSE Then
    strdlgstly = strdlgstly + " | DS_CENTERMOUSE"
    lngValDialog = lngValDialog + DS_CENTERMOUSE
End If
If (DlgBx.lStyle And DS_CONTEXTHELP) = DS_CONTEXTHELP Then
    strdlgstly = strdlgstly + " | DS_CONTEXTHELP"
    lngValDialog = lngValDialog + DS_CONTEXTHELP
End If
If (DlgBx.lStyle And WS_POPUPWINDOW) <> WS_POPUPWINDOW Then
    If (DlgBx.lStyle And WS_POPUP) = WS_POPUP Then
        strdlgstly = strdlgstly + " | WS_POPUP"
        lngValDialog = lngValDialog + WS_POPUP
    End If
    If (DlgBx.lStyle And WS_BORDER) <> WS_BORDER Then
        If (DlgBx.lStyle And WS_DLGFRAME) <> WS_DLGFRAME Then
            If (DlgBx.lStyle And WS_OVERLAPPEDWINDOW) <> WS_OVERLAPPEDWINDOW Then
                If (DlgBx.lStyle And WS_CAPTION) = WS_CAPTION Then
                    strdlgstly = strdlgstly + " | WS_CAPTION"
                    lngValDialog = lngValDialog + WS_CAPTION
                End If
            End If
        Else
            strdlgstly = strdlgstly + " | WS_DLGFRAME"
            lngValDialog = lngValDialog + WS_DLGFRAME
        End If
    ElseIf (DlgBx.lStyle And WS_OVERLAPPEDWINDOW) <> WS_OVERLAPPEDWINDOW Then
        If (DlgBx.lStyle And WS_CAPTION) = WS_CAPTION Then
            strdlgstly = strdlgstly + " | WS_CAPTION"
            lngValDialog = lngValDialog + WS_CAPTION
        Else
            strdlgstly = strdlgstly + " | WS_BORDER"
            lngValDialog = lngValDialog + WS_BORDER
        End If
    End If
    If (DlgBx.lStyle And WS_OVERLAPPEDWINDOW) <> WS_OVERLAPPEDWINDOW Then
        If (DlgBx.lStyle And WS_SYSMENU) = WS_SYSMENU Then
            strdlgstly = strdlgstly + " | WS_SYSMENU"
            lngValDialog = lngValDialog + WS_SYSMENU
        End If
    End If
Else
    strdlgstly = strdlgstly + " | WS_POPUPWINDOW"
    lngValDialog = lngValDialog + WS_POPUPWINDOW
    If (DlgBx.lStyle And WS_DLGFRAME) = WS_DLGFRAME Then
        strdlgstly = strdlgstly + " | WS_DLGFRAME"
        lngValDialog = lngValDialog + WS_DLGFRAME
    End If
End If
If (DlgBx.lStyle And WS_OVERLAPPEDWINDOW) <> WS_OVERLAPPEDWINDOW Then
    If (DlgBx.lStyle And WS_THICKFRAME) = WS_THICKFRAME Then
        strdlgstly = strdlgstly + " | WS_THICKFRAME"
        lngValDialog = lngValDialog + WS_THICKFRAME
    End If
    If (DlgBx.lStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
        strdlgstly = strdlgstly + " | WS_MINIMIZEBOX"
        lngValDialog = lngValDialog + WS_MINIMIZEBOX
    End If
    If (DlgBx.lStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then
        strdlgstly = strdlgstly + " | WS_MAXIMIZEBOX"
        lngValDialog = lngValDialog + WS_MAXIMIZEBOX
    End If
Else
    strdlgstly = strdlgstly + " | WS_OVERLAPPEDWINDOW"
    lngValDialog = lngValDialog + WS_OVERLAPPEDWINDOW
End If

'If (DlgBx.lStyle And WS_OVERLAPPED) = WS_OVERLAPPED Then
'    strdlgstly = strdlgstly + " | WS_OVERLAPPED"
'End If
If (DlgBx.lStyle And WS_CHILD) = WS_CHILD Then
    strdlgstly = strdlgstly + " | WS_CHILD"
    lngValDialog = lngValDialog + WS_CHILD
End If
If (DlgBx.lStyle And WS_MINIMIZE) = WS_MINIMIZE Then
    strdlgstly = strdlgstly + " | WS_MINIMIZE"
    lngValDialog = lngValDialog + WS_MINIMIZE
End If
If (DlgBx.lStyle And WS_VISIBLE) = WS_VISIBLE Then
    strdlgstly = strdlgstly + " | WS_VISIBLE"
    lngValDialog = lngValDialog + WS_VISIBLE
End If
If (DlgBx.lStyle And WS_DISABLED) = WS_DISABLED Then
    strdlgstly = strdlgstly + " | WS_DISABLED"
    lngValDialog = lngValDialog + WS_DISABLED
End If
If (DlgBx.lStyle And WS_CLIPSIBLINGS) = WS_CLIPSIBLINGS Then
    strdlgstly = strdlgstly + " | WS_CLIPSIBLINGS"
    lngValDialog = lngValDialog + WS_CLIPSIBLINGS
End If
If (DlgBx.lStyle And WS_CLIPCHILDREN) = WS_CLIPCHILDREN Then
    strdlgstly = strdlgstly + " | WS_CLIPCHILDREN"
    lngValDialog = lngValDialog + WS_CLIPCHILDREN
End If
If (DlgBx.lStyle And WS_MAXIMIZE) = WS_MAXIMIZE Then
    strdlgstly = strdlgstly + " | WS_MAXIMIZE"
    lngValDialog = lngValDialog + WS_MAXIMIZE
End If
If (DlgBx.lStyle And WS_VSCROLL) = WS_VSCROLL Then
    strdlgstly = strdlgstly + " | WS_VSCROLL"
    lngValDialog = lngValDialog + WS_VSCROLL
End If
If (DlgBx.lStyle And WS_HSCROLL) = WS_HSCROLL Then
    strdlgstly = strdlgstly + " | WS_HSCROLL"
    lngValDialog = lngValDialog + WS_HSCROLL
End If
'If (DlgBx.lStyle And WS_GROUP) = WS_GROUP Then
 '   strdlgstly = strdlgstly + " | WS_GROUP"
'End If
'If (DlgBx.lStyle And WS_TABSTOP) = WS_TABSTOP Then
'    strdlgstly = strdlgstly + " | WS_TABSTOP"
'End If
'If (DlgBx.lStyle And WS_TILEDWINDOW) = WS_TILEDWINDOW Then
'    strdlgstly = strdlgstly + " | WS_TILEDWINDOW"
'End If

'If (DlgBx.lStyle And WS_CHILDWINDOW) = WS_CHILDWINDOW Then
'    strdlgstly = strdlgstly + " | WS_CHILDWINDOW"
'End If
'If (DlgBx.lStyle And WS_EX_DLGMODALFRAME) = WS_EX_DLGMODALFRAME Then
'    strdlgstly = strdlgstly + " | WS_EX_DLGMODALFRAME"
'End If
'If (DlgBx.lStyle And WS_EX_NOPARENTNOTIFY) = WS_EX_NOPARENTNOTIFY Then
'    strdlgstly = strdlgstly + " | WS_EX_NOPARENTNOTIFY"
'End If
'If (DlgBx.lStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST Then
'    strdlgstly = strdlgstly + " | WS_EX_TOPMOST"
'End If
'If (DlgBx.lStyle And WS_EX_ACCEPTFILES) = WS_EX_ACCEPTFILES Then
'    strdlgstly = strdlgstly + " | WS_EX_ACCEPTFILES"
'End If
'If (DlgBx.lStyle And WS_EX_TRANSPARENT) = WS_EX_TRANSPARENT Then
'    strdlgstly = strdlgstly + " | WS_EX_TRANSPARENT"
'End If
If DlgBx.lStyle = lngValDialog Then
    strdlgstly = "STYLE" + Mid(strdlgstly, 3, Len(strdlgstly))
    Debug.Print DlgBx.lStyle
    Debug.Print DlgBx.NumberOfItems
    'Debug.Print lngValDialog
    'Debug.Print strdlgstly
Else
    MsgBox "error please check the DIALOG BOX", vbCritical, "DIALOG BOX ERROR"
End If
    If InStr(1, strdlgstly, "DS_SETFONT", vbTextCompare) > 0 Then
        Dim DlgTmpOff As Long
        Get OpenResFile, Offset + Len(DlgBx) + 4 + (IntTmp * 2), DlgbxFontSize
        Get OpenResFile, Offset + Len(DlgBx) + 4 + (IntTmp * 2) + 2, StrTmp
        DlgTmpOff = Offset + Len(DlgBx) + 4 + (IntTmp * 2) + 2
        IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar, vbTextCompare)
        DlgTmpOff = (DlgTmpOff + IntTmp * 2) + 2
        DlgBxFontName = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp)
    End If
    Dim DlgBxTmp() As DLGITEMTEMPLATE
    ReDim DlgBxTmp(1 To DlgBx.NumberOfItems)
    Get OpenResFile, Offset + Len(DlgBx) + 4 + (IntTmp * 2), DlgbxFontSize
    MsgBox Offset + Len(DlgBx) + 4 + (IntTmp * 2)
    MsgBox Offset + Len(DlgBx) + 4 + (IntTmp * 2) + 2
    For i = 1 To DlgBx.NumberOfItems
        Get OpenResFile, DlgTmpOff, DlgBxTmp
        Debug.Print DlgBxTmp(i).style
        Debug.Print DlgBxTmp(i).dwExtendedStyle
        Debug.Print DlgBxTmp(i).x
        Debug.Print DlgBxTmp(i).y
        Debug.Print DlgBxTmp(i).cx
        Debug.Print DlgBxTmp(i).cy
        Get OpenResFile, DlgTmpOff + Len(DlgBxTmp(i)) + 4, StrTmp
        IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar, vbTextCompare)
        DlgBxFontName = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp)
        DlgTmpOff = DlgTmpOff + Len(DlgBxTmp(i)) + (IntTmp * 2)
    Next i
        
End If
End Function
Private Function DumpMessageTable(Offset As Long, MessageTableSize As Long, OpenResFile As Long)
Dim MsgTableData As MESSAGERESOURCEDATA
Dim MsgTableBlock As MESSAGERESOURCEBLOCK
Dim MsgTableEntry As MESSAGERESOURCEENTRY
Dim MsgData As String
Dim freenum As Integer

freenum = FreeFile
Open "c:\6110\" & BinName & "." & "messagetable.txt" For Binary As freenum
Dim BlockLoop As Long, BlockOffset As Long, EntryLoop As Long, EntryOffset As Long

    Get OpenResFile, Offset, MsgTableData
    Put freenum, LOF(freenum) + 1, "Number of Blocks - " & MsgTableData.NumberOfBlocks & vbNewLine
    For BlockLoop = 1 To MsgTableData.NumberOfBlocks
        Put freenum, , " BLOCK - " & BlockLoop & vbNewLine
        Get OpenResFile, Offset + Len(MsgTableData) + BlockOffset, MsgTableBlock
        EntryOffset = MsgTableBlock.OffsetToEntries
        Put freenum, , "    Low ID   - " & MsgTableBlock.LowId & vbNewLine
        Put freenum, , "    HighID   - " & MsgTableBlock.HighId & vbNewLine
        Put freenum, , "    Offset to Entries - " & CLng(MsgTableBlock.OffsetToEntries) & vbNewLine
        For EntryLoop = 0 To MsgTableBlock.HighId - MsgTableBlock.LowId
            Get OpenResFile, Offset + EntryOffset, MsgTableEntry
            MsgData = String(MsgTableEntry.Length, " ")
            Get OpenResFile, Offset + EntryOffset + 4, MsgData
            EntryOffset = EntryOffset + MsgTableEntry.Length
            Put freenum, , "        ID - " & MsgTableBlock.LowId + EntryLoop & "..." & MsgTableBlock.LowId + EntryLoop & vbNewLine
            Put freenum, , "        Length - " & MsgTableEntry.Length & vbNewLine
            Put freenum, , "        Flags  - " & MsgTableEntry.flags & vbNewLine
            If MsgTableEntry.flags = 0 Then
                Put freenum, , "        Text - " & MsgData & vbNewLine
            Else
                Put freenum, , "        Text - " & StrConv(MsgData, vbFromUnicode) & vbNewLine
            End If
        Next EntryLoop
        BlockOffset = BlockOffset + Len(MsgTableBlock)
    Next BlockLoop
End Function
Private Function DumpAnyResource(Offset As Long, ResName As String, ResSize As Long, OpenResFile As Long)
Dim freenum As Integer
Dim StrDump As String
freenum = FreeFile
Open "c:\6110\" & BinName & "." & ResName For Binary As freenum
StrDump = String(ResSize, " ")
Get OpenResFile, Offset, StrDump
Put freenum, , StrDump
Close freenum
End Function
Private Function DumpMenuTable(Offset As Long, ResName As String, ResSize As Long, OpenResFile As Long)
Dim intVer As Integer
Get OpenResFile, Offset, intVer
Dim freenum As Integer
freenum = FreeFile
Open "c:\6110\" & BinName & "." & "menu.rc" For Binary As freenum
Put freenum, LOF(freenum) + 1, ResName & " Menu" & vbNewLine
Put freenum, , "Begin" & vbNewLine
If intVer = 1 Then
'                        Dim MenuExHeader As MENUEXTEMPLATEHEADER
'                        Dim menuexitem As MENUEXTEMPLATEITEM
'                        Get openresfile, offset, MenuExHeader
'                        MsgBox MenuExHeader.wOffset
'                        Get openresfile, offset + Len(MenuExHeader) + MenuExHeader.wOffset, menuexitem
'                        Debug.Print menuexitem.dwType
'                        Debug.Print menuexitem.szText
'                        Debug.Print menuexitem.uId
'                        Debug.Print menuexitem.bResInfo
'                        Debug.Print menuexitem.dwState
'                        Get openresfile, offset + Len(MenuExHeader) + MenuExHeader.wOffset + Len(menuexitem), menuexitem
'                        MsgBox offset + Len(MenuExHeader) + MenuExHeader.wOffset + Len(menuexitem) + Len(menuexitem)
'                        MsgBox menuexitem.bResInfo
Else
    Dim MenuHeader As MenuHeader
    Dim MeNoIt As NORMALMENUITEM
    Dim MePoIt As POPUPMENUITEM
    Dim StrTmp As String
    Dim IntTmp As Integer
    Dim MeCap As String
    Dim Popcap As String
    Dim popOff As Long
    Dim ReLoopEr As Boolean
    Dim IntPop As Integer
    StrTmp = String(255, " ")
    Get OpenResFile, Offset, MenuHeader
   popOff = Len(MenuHeader)
getPop:
    Do
        Get OpenResFile, Offset + popOff, MeNoIt
        If (MeNoIt.resInfo And MF_POPUP) <> MF_POPUP Then
            Get OpenResFile, Offset + popOff, MePoIt
            Get OpenResFile, Offset + popOff + Len(MePoIt), StrTmp
            IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar)
            Popcap = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp - 1)
            If Popcap = "" Then
                Put freenum, , "        MENUITEM SEPARATOR" & vbNewLine
            Else
                Put freenum, , "        MENUITEM " & """" & Popcap & """, " & MePoIt.PopId & vbNewLine
            End If
            popOff = popOff + (IntTmp * 2) + Len(MePoIt)
            Debug.Print popOff
            If (MePoIt.resInfo And MF_END) = MF_END Then
                Put freenum, , "    END" & vbNewLine
                IntPop = IntPop - 1
                If popOff = ResSize Then
                    Put freenum, , "    END" & vbNewLine
                    GoTo ReLoop
                End If
            End If
            MePoIt.resInfo = 1
            GoTo getPop
        End If
        Get OpenResFile, Offset + popOff + 2, StrTmp
        IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar)
        MeCap = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp - 1)
        If MeCap = "" And (MeNoIt.resInfo And MF_POPUP) <> MF_POPUP Then
            If (MeNoIt.resInfo And MF_END) = MF_END Then
                IntPop = IntPop - 1
            End If
            Put freenum, , "    MENUITEM SEPARATOR" & vbNewLine
            popOff = popOff + (IntTmp * 2) + Len(MePoIt)
            Debug.Print popOff
            GoTo ReLoop
        End If
        If (MeNoIt.resInfo And MF_POPUP) <> MF_POPUP Then
            If (MeNoIt.resInfo And MF_END) = MF_END Then
                IntPop = IntPop - 1
            End If
            Put freenum, , "    MENUITEM " & """" & MeCap & """" & vbNewLine
            Debug.Print MeCap
            popOff = popOff + (IntTmp * 2) + Len(MeNoIt)
            Debug.Print popOff
            MeNoIt.resInfo = 1
            GoTo ReLoop
        End If
        Put freenum, , "    Popup " & """" & MeCap & """" & vbNewLine
        Debug.Print MeCap
        IntPop = IntPop + 1
        Put freenum, , "    BEGIN" & vbNewLine
        popOff = popOff + (IntTmp * 2) + 2
        Debug.Print popOff
        Do While (MePoIt.resInfo And MF_END) <> MF_END
            Get OpenResFile, Offset + popOff, MePoIt
            If (MePoIt.resInfo And MF_POPUP) = MF_POPUP Then
                MePoIt.resInfo = 1
                'IntPop = IntPop + 1
                GoTo ReLoop
            End If
            Get OpenResFile, Offset + popOff + Len(MePoIt), StrTmp
            IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar)
            Popcap = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp - 1)
            If Popcap = "" Then
                Put freenum, , "        MENUITEM SEPARATOR" & vbNewLine
            Else
                Put freenum, , "        MENUITEM " & """" & Popcap & """, " & MePoIt.PopId & vbNewLine
            End If
            popOff = popOff + (IntTmp * 2) + Len(MePoIt)
            Debug.Print popOff
            If (MePoIt.resInfo And MF_END) = MF_END Then
                IntPop = IntPop - 1
                Put freenum, , "    END" & vbNewLine
            End If
            If IntPop >= 1 And (MeNoIt.resInfo And MF_END) = MF_END Then
                IntPop = IntPop - 1
                MeNoIt.resInfo = 1
            End If
        Loop
        MePoIt.resInfo = 1
        If (MeNoIt.resInfo And MF_END) = MF_END Or MeNoIt.resInfo = 1 Then
            'IntPop = IntPop - 1
            Put freenum, , "    END" & vbNewLine
        End If
ReLoop:
Debug.Print popOff
    Loop While (MeNoIt.resInfo And MF_END) <> MF_END Or IntPop > 0
Put freenum, , "//==============================//" & vbNewLine
End If
End Function
Private Function DumpMenuTable2(Offset As Long, ResName As String, ResSize As Long, OpenResFile As Long)
Dim intVer As Integer
Get OpenResFile, Offset, intVer
Dim freenum As Integer
freenum = FreeFile
Open "c:\6110\" & BinName & "." & "menu.rc" For Binary As freenum
Put freenum, , "//==============================//" & vbNewLine
Put freenum, LOF(freenum) + 1, ResName & " Menu" & vbNewLine
Put freenum, , "Begin" & vbNewLine
If intVer = 1 Then

Else
    Dim MenuHeader As MenuHeader
    Dim MeNoIt As NORMALMENUITEM
    Dim MePoIt As POPUPMENUITEM
    Dim StrTmp As String
    Dim IntTmp As Integer
    Dim MeCap As String
    Dim Popcap As String
    Dim popOff As Long
    Dim LastPop As Integer
    Dim BeginPop As Integer
    Dim EndPop As Integer
    Dim CurPop As Boolean
    StrTmp = String(255, " ")
    BeginPop = 1
    Get OpenResFile, Offset, MenuHeader
    popOff = Len(MenuHeader)
    Do While popOff < ResSize
        Get OpenResFile, Offset + popOff, MeNoIt
        Select Case True
            Case (MeNoIt.resInfo And MF_POPUP) = MF_POPUP
                If (MeNoIt.resInfo And MF_END) = MF_END Then
                    LastPop = LastPop + 1
                    CurPop = True
                Else
                    CurPop = False
                End If
                BeginPop = BeginPop + 1
                Get OpenResFile, Offset + popOff + 2, StrTmp
                IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar)
                MeCap = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp - 1)
                popOff = popOff + (IntTmp * 2) + Len(MeNoIt)
                Put freenum, , "    Popup " & """" & MeCap & """" & vbNewLine
                Put freenum, , "    BEGIN " & vbNewLine
                If (MeNoIt.resInfo And MF_END) = MF_END Then
                    LastPop = LastPop + 1
                End If
            Case (MeNoIt.resInfo And MF_POPUP) <> MF_POPUP
                If MeNoIt.resInfo = 0 Or MeNoIt.resInfo = 128 Then
                Else
                    'MsgBox "MENU, PLEASE CHECK THIS = :" & MeNoIt.resInfo, vbCritical, "MENU"
                End If
                Get OpenResFile, Offset + popOff, MePoIt
                Get OpenResFile, Offset + popOff + Len(MePoIt), StrTmp
                IntTmp = InStr(1, StrConv(StrTmp, vbFromUnicode), vbNullChar)
                Popcap = Mid(StrConv(StrTmp, vbFromUnicode), 1, IntTmp - 1)
                If Popcap = "" Then
                        Put freenum, , "        MENUITEM SEPARATOR" & vbNewLine
                Else
                    Put freenum, , "        MENUITEM " & """" & Popcap & """, " & MePoIt.PopId & vbNewLine
                End If
                popOff = popOff + (IntTmp * 2) + Len(MePoIt)
                If (MePoIt.resInfo And MF_END) = MF_END Then
                    Put freenum, , "    END" & vbNewLine
                    EndPop = EndPop + 1
                    If CurPop = True Then
                        LastPop = LastPop - 1
                        Put freenum, , "    END" & vbNewLine
                        EndPop = EndPop + 1
                        CurPop = False
                    End If
                End If
        End Select
    Loop
    For EndPop = EndPop To BeginPop - 1
           Put freenum, , "    END" & vbNewLine
    Next EndPop
    Put freenum, , "//==============================//" & vbNewLine
End If
Close freenum
End Function
Sub Main()
    GetPEinfo
End Sub
