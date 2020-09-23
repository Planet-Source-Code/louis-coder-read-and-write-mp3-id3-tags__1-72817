Attribute VB_Name = "Stuffmod"
Option Explicit
'(c)2002, 2003, 2004 by Louis. Contains often used subs/functions.
'
'NOTE: this module may grow permanently. Add any piece of code
'whenever you think it could serve as basic function in further programs.
'Use the prefix 'Stuff_' to avoid name conflics with procedures in other modules.
'
'Stuff_GFCDGetFileName
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Stuff_GFCDSetFileName
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Stuff_GFCDGetColor
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORSTRUCT) As Long
'Stuff_GFSelectDirectory
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'StuffFastLine_Draw
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'StuffFastLine_SetColor
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Stuff_GetStringLong
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'STUFF_COPYMEMORY
Public Declare Sub STUFF_COPYMEMORY Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'STUFF_BITBLT
Public Declare Function STUFF_BITBLT Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'STUFF_SLEEP
Public Declare Sub STUFF_SLEEP Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
'StuffFastLine_SetColor
Public Const PS_SOLID As Long = 0
'Stuff_GFCDGetFileName; Stuff_GFCDSetFileName
Private Const OFN_HIDEREADONLY = &H4
Dim NULLARRAYSTRING(0 To 0) As String 'disable if already existing in target project
'Stuff_GFCDGetFileName; Stuff_GFCDSetFileName
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'Stuff_GFCDGetColor
Private Type CHOOSECOLORSTRUCT
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'Stuff_GFSelectDirectory
Private Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type
'Stuff_GFCDGetColor
Private Const CC_RGBINIT = &H1
'Stuff_GFSelectDirectory
'Const STUFF_MAX_PATH = 260
Private Const ERROR_SUCCESS As Long = 0
Private Const CSIDL_DESKTOP As Long = &H0
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_STATUSTEXT As Long = &H4
Private Const BIF_RETURNFSANCESTORS As Long = &H8
'StuffFastLine
Dim StuffFastLinePenHandleUnchanged As Long

'***CHOOSING FUNCTIONS***

Public Function Stuff_GFCDGetFileNameFast(ByVal Title As String, ByVal DefaultPath As String) As String
    'on error resume next 'requests a file name, file of type *.*
    Dim FilterNumber As Integer
    Dim FilterDescriptionArray(1 To 1) As String
    Dim FilterStringArray(1 To 1) As String
    'preset
    FilterNumber = 1
    FilterDescriptionArray(1) = "All Files"
    FilterStringArray(1) = "*.*"
    'begin
    Stuff_GFCDGetFileNameFast = Stuff_GFCDGetFileName(Title, FilterNumber, FilterDescriptionArray(), FilterStringArray(), 1, DefaultPath)
End Function

Public Function Stuff_GFCDGetFileName(ByVal Title As String, ByRef FilterNumber As Integer, ByRef FilterDescriptionArray() As String, ByRef FilterStringArray() As String, ByVal DefaultFilterIndex As Integer, ByVal DefaultPath As String) As String
    'on error resume next 'must be placed into a form (uses hWnd); FilterNumber may be 0 (then pass NULLARRAYSTRING()), DefaultPath should contain path and file name
    Dim OPENFILENAMEVar As OPENFILENAME
    Dim DefaultFileName As String
    Dim DefaultDirectoryName As String
    Dim Temp As Long
    '
    'NOTE: the FilerDescriptionArray() and the FilterStringArray() data
    'must have the following format (example; description/string):
    '
    'Bitmap/*.bmp;*.jpg;*.gif
    '
    'If FilterNumber is 0, the preset filter 'All Files/*.*' is used.
    'If the user pressed 'Cancel' the function returns nothing ("").
    '
    'initialize structure
    OPENFILENAMEVar.lStructSize = Len(OPENFILENAMEVar)
    OPENFILENAMEVar.hWndOwner = 0 'do not use form (module ?!) handle
    OPENFILENAMEVar.hInstance = App.hInstance
    If Not (FilterNumber = 0) Then
        '
        'NOTE: the filter string contains string pairs (filter description/string),
        'the string end is marked by two null chars.
        '
        For Temp = 1 To FilterNumber
            OPENFILENAMEVar.lpstrFilter = OPENFILENAMEVar.lpstrFilter + _
                FilterDescriptionArray(Temp) + Chr$(0) + FilterStringArray(Temp) + Chr$(0)
        Next Temp
        OPENFILENAMEVar.lpstrFilter = OPENFILENAMEVar.lpstrFilter + Chr$(0) 'two null chars at filter string end
    Else
        OPENFILENAMEVar.lpstrFilter = "All Files" + Chr$(0) + "*.*" + Chr$(0) + Chr$(0)
    End If
    OPENFILENAMEVar.nFilterIndex = DefaultFilterIndex
    If Not (Stuff_GetFileName(DefaultPath) = "") Then
        DefaultFileName = Stuff_GetFileName(DefaultPath)
        OPENFILENAMEVar.nMaxFile = 260 + 1 'STUFF_MAX_PATH
        OPENFILENAMEVar.lpstrFile = String$(260 + 1, Chr$(0))
        Mid$(OPENFILENAMEVar.lpstrFile, 1, Len(DefaultFileName)) = Left$(DefaultFileName, 260)
    Else
        OPENFILENAMEVar.nMaxFile = 260 + 1 'STUFF_MAX_PATH
        OPENFILENAMEVar.lpstrFile = String$(260 + 1, Chr$(0))
    End If
    OPENFILENAMEVar.lpstrTitle = Title + Chr$(0)
    DefaultDirectoryName = Left$(DefaultPath, Len(DefaultPath) - Len(DefaultFileName))
    OPENFILENAMEVar.lpstrInitialDir = DefaultDirectoryName + Chr$(0)
    OPENFILENAMEVar.nMaxFileTitle = 260 + 1 'STUFF_MAX_PATH
    OPENFILENAMEVar.lpstrFileTitle = String$(260 + 1, Chr$(0)) 'receives selected file name (without directory)
    OPENFILENAMEVar.Flags = OFN_HIDEREADONLY
    'end of initializing structure
    If Not (GetOpenFileName(OPENFILENAMEVar) = 0) Then
        If Not (InStr(1, OPENFILENAMEVar.lpstrFile, Chr$(0), vbBinaryCompare) = 0) Then 'verify
            Stuff_GFCDGetFileName = Left$(OPENFILENAMEVar.lpstrFile, InStr(1, OPENFILENAMEVar.lpstrFile, Chr$(0), vbBinaryCompare) - 1)
        Else
            Stuff_GFCDGetFileName = OPENFILENAMEVar.lpstrFile
        End If
    Else
        Stuff_GFCDGetFileName = "" 'reset (error)
    End If
End Function

Public Function Stuff_GFCDSetFileName(ByVal Title As String, ByRef FilterNumber As Integer, ByRef FilterDescriptionArray() As String, ByRef FilterStringArray() As String, ByVal DefaultFilterIndex As Integer, ByVal DefaultPath As String) As String
    'on error resume next 'must be placed into a form (uses hWnd); FilterNumber may be 0 (then pass NULLARRAYSTRING()), DefaultPath should contain path and file name
    Dim OPENFILENAMEVar As OPENFILENAME
    Dim DefaultFileName As String
    Dim DefaultDirectoryName As String
    Dim Temp As Long
    '
    'NOTE: the FilerDescriptionArray() and the FilterStringArray() data
    'must have the following format (example; description/string):
    '
    'Bitmap/*.bmp;*.jpg;*.gif
    '
    'If FilterNumber is 0, the preset filter 'All Files/*.*' is used.
    'If the user pressed 'Cancel' the function returns nothing ("").
    '
    'initialize structure
    OPENFILENAMEVar.lStructSize = Len(OPENFILENAMEVar)
    OPENFILENAMEVar.hWndOwner = 0 'do not use form (module ?!) handle
    OPENFILENAMEVar.hInstance = App.hInstance
    If Not (FilterNumber = 0) Then
        '
        'NOTE: the filter string contains string pairs (filter description/string),
        'the string end is marked by two null chars.
        '
        For Temp = 1 To FilterNumber
            OPENFILENAMEVar.lpstrFilter = OPENFILENAMEVar.lpstrFilter + _
                FilterDescriptionArray(Temp) + Chr$(0) + FilterStringArray(Temp) + Chr$(0)
        Next Temp
        OPENFILENAMEVar.lpstrFilter = OPENFILENAMEVar.lpstrFilter + Chr$(0) 'two null chars at filter string end
    Else
        OPENFILENAMEVar.lpstrFilter = "All Files" + Chr$(0) + "*.*" + Chr$(0) + Chr$(0)
    End If
    OPENFILENAMEVar.nFilterIndex = DefaultFilterIndex
    If Not (Stuff_GetFileName(DefaultPath) = "") Then
        DefaultFileName = Stuff_GetFileName(DefaultPath)
        OPENFILENAMEVar.nMaxFile = 260 + 1 'STUFF_MAX_PATH
        OPENFILENAMEVar.lpstrFile = String$(260 + 1, Chr$(0))
        Mid$(OPENFILENAMEVar.lpstrFile, 1, Len(DefaultFileName)) = Left$(DefaultFileName, 260)
    Else
        OPENFILENAMEVar.nMaxFile = 260 + 1 'STUFF_MAX_PATH
        OPENFILENAMEVar.lpstrFile = String$(260 + 1, Chr$(0))
    End If
    OPENFILENAMEVar.lpstrTitle = Title + Chr$(0)
    DefaultDirectoryName = Left$(DefaultPath, Len(DefaultPath) - Len(DefaultFileName))
    OPENFILENAMEVar.lpstrInitialDir = DefaultDirectoryName + Chr$(0)
    OPENFILENAMEVar.nMaxFileTitle = 260 + 1 'STUFF_MAX_PATH
    OPENFILENAMEVar.lpstrFileTitle = String$(260 + 1, Chr$(0)) 'receives selected file name (without directory)
    OPENFILENAMEVar.Flags = OFN_HIDEREADONLY
    'end of initializing structure
    If Not (GetSaveFileName(OPENFILENAMEVar) = 0) Then
        If Not (InStr(1, OPENFILENAMEVar.lpstrFile, Chr$(0), vbBinaryCompare) = 0) Then 'verify
            Stuff_GFCDSetFileName = Left$(OPENFILENAMEVar.lpstrFile, InStr(1, OPENFILENAMEVar.lpstrFile, Chr$(0), vbBinaryCompare) - 1)
        Else
            Stuff_GFCDSetFileName = OPENFILENAMEVar.lpstrFile
        End If
    Else
        Stuff_GFCDSetFileName = "" 'reset (error)
    End If
End Function

Public Function Stuff_GFCDGetColor(ByVal DefaultColor As Long, ByVal UserColorNumberPassed As Integer, ByRef UserColorArrayPassed() As Long) As Long
    'on error resume next 'v1.0 (no user color support); returns True if user aborted (always check if return value is True)
    Dim CHOOSECOLORSTRUCTVar As CHOOSECOLORSTRUCT
    Dim UserColorArray(1 To 16) As Long
    Dim UserColorLoop As Integer
    '
    'NOTE: the ChooseColor function requires to be able to
    'access an array of exactly 16 COLORREF (Long) variables, otherwise
    'the program will crash.
    'The passed array must not contain 16 values, this function
    'will transfer the user color values of the passed user color array.
    '
    'preset
    For UserColorLoop = 1 To UserColorNumberPassed
        If UserColorLoop > 16 Then Exit For
        UserColorArray(UserColorLoop) = UserColorArrayPassed(UserColorLoop)
    Next UserColorLoop
    'initialize structure
    CHOOSECOLORSTRUCTVar.lStructSize = Len(CHOOSECOLORSTRUCTVar)
    CHOOSECOLORSTRUCTVar.hWndOwner = 0 'do not use an owner window (module?)
    CHOOSECOLORSTRUCTVar.hInstance = App.hInstance
    CHOOSECOLORSTRUCTVar.rgbResult = DefaultColor
    CHOOSECOLORSTRUCTVar.Flags = CC_RGBINIT
    CHOOSECOLORSTRUCTVar.lpCustColors = VarPtr(UserColorArray(1))
    CHOOSECOLORSTRUCTVar.lCustData = 0
    CHOOSECOLORSTRUCTVar.lpfnHook = 0
    CHOOSECOLORSTRUCTVar.lpTemplateName = Chr$(0)
    'end of initializing structure
    'begin
    If Not (ChooseColor(CHOOSECOLORSTRUCTVar) = 0) Then 'verify
        Stuff_GFCDGetColor = CHOOSECOLORSTRUCTVar.rgbResult 'ok
    Else
        Stuff_GFCDGetColor = True 'error
    End If
End Function

Public Function Stuff_GFSelectDirectory(ByVal RootDirectory As String, ByVal InfoText As String) As String
    'on error Resume Next 'v1.0 - does not support a root directory
    Dim BROWSEINFOVar As BROWSEINFO
    Dim Temp As Long
    Dim Tempstr$
    'preset
    'BROWSEINFOVar.pidlRoot = RootDirectory 'does not work
    BROWSEINFOVar.hOwner = 0 'Me.hWnd
    BROWSEINFOVar.pszDisplayName = String$(260, Chr$(0)) 'display name (i.e. 'Windows' for C:\Windows\)
    BROWSEINFOVar.lpszTitle = InfoText
    BROWSEINFOVar.ulFlags = BIF_RETURNONLYFSDIRS 'file system directories only
    BROWSEINFOVar.lpfn = 0 'address of event call-back function
    BROWSEINFOVar.lParam = 0 'parameter that would be passed to event call-back function
    'begin
    Temp = SHBrowseForFolder(BROWSEINFOVar)
    'return selected folder
    'BROWSEINFOVar.pszDisplayName 'display name of selected folder
    'BROWSEINFOVar.iImage 'image of selected item in system image list
    If Not (Temp = 0) Then 'verify
        Tempstr$ = String$(260, Chr$(0))
        Call SHGetPathFromIDList(ByVal Temp, ByVal Tempstr$)
        If Not (InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) = 0) Then 'verify
            Stuff_GFSelectDirectory = Left$(Tempstr$, InStr(1, Tempstr$, Chr$(0), vbBinaryCompare) - 1) 'ok
        Else
            Stuff_GFSelectDirectory = "" 'error
        End If
    Else
        Stuff_GFSelectDirectory = "" 'error
    End If
End Function

'***END OF CHOOSING FUNCTIONS***
'***FILE SYSTEM FUNCTIONS***

Public Function Stuff_GetExtendedFileName(ByVal FileMainName As String, ByVal FileNameExtension As String, FileMainNameSuffix As String) As String 'general function (may be used in any project)
    On Error GoTo Error: 'important; v1.1
    Stuff_GetExtendedFileName = "" 'reset
    '
    'NOTE: example: passing ("C:\VisualBasic", ".EXE", "#") will return the following
    'strings (depending on number of files already created using this function):
    '
    '"C:\VisualBasic.EXE"
    '"C:\VisualBasic#2.EXE"
    '[...]
    '"C:\VisualBasic#256.EXE"
    '""
    '
    If Not (FileMainName + FileNameExtension = "") Then
        If (Dir$(FileMainName + FileNameExtension) = "") And (Dir$(FileMainName + FileMainNameSuffix + LTrim$(Str$(1)) + FileNameExtension) = "") Then
            Stuff_GetExtendedFileName = FileMainName + FileNameExtension
            Exit Function
        End If
        Dim Temp As Long
        For Temp = 2 To 256
            If Dir$(FileMainName + FileMainNameSuffix + LTrim$(Str$(Temp)) + FileNameExtension) = "" Then
                Stuff_GetExtendedFileName = FileMainName + FileMainNameSuffix + LTrim$(Str$(Temp)) + FileNameExtension
                Exit Function
            End If
        Next Temp
    End If
    Exit Function
Error:
    Stuff_GetExtendedFileName = "" 'reset (error)
    Exit Function
End Function

Public Function Stuff_GenerateTempFileName(ByVal TempFilePath As String) As String
    'on error Resume Next 'returns name of a not-existing file in TempFilePath, file name has following format: ########.tmp
    Dim Temp As Integer
    '
    'NOTE: this function (26.12.2002) uses DirSave() instead of the VB Dir$().
    '
    'begin
    If (Not (Right$(TempFilePath, 1) = "\")) And (Not (TempFilePath = "")) Then
        TempFilePath = TempFilePath + "\"
    End If
    Do
        Stuff_GenerateTempFileName = TempFilePath + Format$((Rnd(1) * 1E+08!), "00000000") + ".tmp"
        Temp = Temp + 1 'save is save
    Loop Until (DirSave(Stuff_GenerateTempFileName) = "") Or (Temp = 32767)
End Function

Public Function Stuff_GetFileName(ByVal GetFileNameName As String) As String 'also used by Hmod.KeyHook_Open()
    'on error Resume Next 'returns chars after last backslash or nothing
    Dim GetFileNameLoop As Integer
    Stuff_GetFileName = "" 'reset
    For GetFileNameLoop = Len(GetFileNameName) To 1 Step (-1)
        If Mid$(GetFileNameName, GetFileNameLoop, 1) = "\" Then
            Stuff_GetFileName = Right$(GetFileNameName, Len(GetFileNameName) - GetFileNameLoop)
            Exit For
        End If
    Next GetFileNameLoop
End Function

Public Function Stuff_GetFileMainName(ByVal File As String) As String
    'on error Resume Next 'return chars before last "." or File
    Dim GetFileMainNameLoop As Long
    Stuff_GetFileMainName = File 'preset
    For GetFileMainNameLoop = Len(File) To 1 Step (-1)
        If Mid$(File, GetFileMainNameLoop, 1) = "." Then
            Stuff_GetFileMainName = Left$(File, GetFileMainNameLoop - 1)
            Exit For
        End If
    Next GetFileMainNameLoop
End Function

Public Function Stuff_GetFileNameSuffix(ByVal File As String) As String
    'on error Resume Next 'return chars after last "." or nothing
    Dim GetFileNameSuffixLoop As Long
    Stuff_GetFileNameSuffix = "" 'reset
    For GetFileNameSuffixLoop = Len(File) To 1 Step (-1)
        If Mid$(File, GetFileNameSuffixLoop, 1) = "." Then
            Stuff_GetFileNameSuffix = Right$(File, Len(File) - GetFileNameSuffixLoop)
            Exit For
        End If
    Next GetFileNameSuffixLoop
End Function

Public Function Stuff_IsFullPath(ByVal File As String) As Boolean
    'on error resume next 'something new since MP3 Renamer 2, returns True if File is a full path, False if not
    '
    'NOTE: to be a full path File must contain one directory- and one file name.
    '
    If (InStr(1, File, "\", vbBinaryCompare)) Then 'check first to increase speed
        If Stuff_GetDirectoryName(File) = "" Then GoTo Error:
        If Stuff_GetFileName(File) = "" Then GoTo Error:
        Stuff_IsFullPath = True 'ok
        Exit Function
    Else
        GoTo Error:
    End If
    Exit Function
Error:
    Stuff_IsFullPath = False 'error
    Exit Function
End Function

Public Function Stuff_IsFileExisting(ByVal File As String) As Boolean 'GFSkinEngine specific
    'on error resume next 'returns True if passed File (full path) is existing, False if not
    '
    'NOTE: this function was implemented for compatibility with Stuff_IsFullPath()
    'only, still use the 'conventional' checking method (see below).
    '
    If (Len(File)) Then 'check first to increase speed
        Stuff_IsFileExisting = Not ((DirSave(File) = "") Or (Right$(File, 1) = "\")) 'Len() = 0 already checked
        Exit Function
    Else
        Stuff_IsFileExisting = False 'error
        Exit Function
    End If
End Function

Public Function Stuff_IsFileCreatable(ByVal File As String) As Boolean
    On Error GoTo Error: 'important; checks if a file could be created (overwritten), does not touch any existing file; pass a full path including file name
    Dim FreeFileNumber As Integer
    '
    'NOTE: you can use this function to quickly check if a file could be written
    '(if an output path or directory is valid):
    'IsPathValidFlag = Stuff_IsFileCreatable(OutputDirectory + "Test.dat").
    '
    'begin
    If (Right$(File, 1) = "\") Or (Len(File) = 0) Then
        Stuff_IsFileCreatable = False
        Exit Function
    End If
    If Stuff_IsFileExisting(File) = True Then
        If ((GetAttr(File) And vbReadOnly) = 0) And ((GetAttr(File) And vbHidden) = 0) And ((GetAttr(File) And vbSystem) = 0) Then 'verify
            Stuff_IsFileCreatable = True
        Else
            Stuff_IsFileCreatable = False
        End If
    Else
        FreeFileNumber = FreeFile(0)
        Open File For Output As #FreeFileNumber 'jump to Error: if not possible
        Close #FreeFileNumber
        If Stuff_IsFileExisting(File) Then Kill File 'if creatable then also deletable
        Stuff_IsFileCreatable = True
    End If
    Exit Function
Error:
    Close #FreeFileNumber 'make sure file is closed
    If Stuff_IsFileExisting(File) Then Kill File 'make sure file is deleted
    Stuff_IsFileCreatable = False
    Exit Function
End Function

Public Function Stuff_GetDirectoryName(ByVal GetDirectoryNameName As String) As String
    'on error Resume Next 'returns chars from string begin to (including) last backslash or nothing
    Dim GetDirectoryNameLoop As Integer
    Stuff_GetDirectoryName = "" 'reset
    For GetDirectoryNameLoop = Len(GetDirectoryNameName) To 1 Step (-1)
        If Mid$(GetDirectoryNameName, GetDirectoryNameLoop, 1) = "\" Then
            Stuff_GetDirectoryName = Left$(GetDirectoryNameName, GetDirectoryNameLoop)
            Exit For
        End If
    Next GetDirectoryNameLoop
End Function

Public Function Stuff_GetRootDir(ByVal GetRootDirPath As String) As String
    'on error Resume Next 'returns root dir of passed path, even if located on a network machine
    Dim GetRootDirLoop As Integer
    'verify
    GetRootDirPath = Left$(GetRootDirPath, 32767)
    'begin
    If Not (Left$(GetRootDirPath, 2) = "\\") Then
        Stuff_GetRootDir = Left$(GetRootDirPath, 3) 'i.e. c:\
    Else
        Stuff_GetRootDir = Chr$(0) 'preset (error)
        GetRootDirPath = GetRootDirPath + "\" 'add end sign (testing is not required, increase speed)
        For GetRootDirLoop = 3 To Len(GetRootDirPath)
            If Mid$(GetRootDirPath, GetRootDirLoop, 1) = "\" Then
                Select Case Stuff_GetRootDir
                Case Chr$(0)
                    Stuff_GetRootDir = ""
                Case ""
                    Stuff_GetRootDir = Left$(GetRootDirPath, GetRootDirLoop) 'i.e. \\SERVER\C\
                    Exit For
                End Select
            End If
        Next GetRootDirLoop
        If Stuff_GetRootDir = Chr$(0) Then Stuff_GetRootDir = "" 'reset (error)
    End If
End Function

'PRIVATE:

Private Function DirSave(ByRef PathName As String, Optional ByVal Attributes As Integer = vbNormal) As String
    On Error GoTo Error: 'important
    '
    'NOTE: Dir$() raises an error if PathName represents a cdrom drive
    'and the cd is not inserted (damn VB!). Use this function rather than Dir$().
    '
    DirSave = Dir$(PathName, Attributes) 'ok
    Exit Function
Error:
    DirSave = "" 'error
    Exit Function
End Function

'END OF PRIVATE

'***END OF FILE SYSTEM FUNCTIONS***
'***CONVERSION FUNCTIONS***

Public Function Stuff_GetLongString(ByVal LongValue As Long) As String
    'on error resume next 'get the 4 bytes of a Long value
    Stuff_GetLongString = String$(4, Chr$(0))
    Call CopyMemory(ByVal Stuff_GetLongString, LongValue, 4)
End Function

Public Function Stuff_GetStringLong(ByVal StringString As String) As Long
    'on error resume next
    Call CopyMemory(Stuff_GetStringLong, ByVal StringString, 4)
End Function

Public Function Stuff_GetDoubleString(ByVal DoubleValue As Double) As String
    'on error resume next 'get the 8 bytes of a Double value
    Stuff_GetDoubleString = String$(8, Chr$(0))
    Call CopyMemory(ByVal Stuff_GetDoubleString, DoubleValue, 8)
End Function

Public Function Stuff_GetStringDouble(ByVal StringString As String) As Double
    'on error resume next
    Call CopyMemory(Stuff_GetStringDouble, ByVal StringString, 8)
End Function

Public Function Stuff_GetBooleanString(ByVal BooleanValue As Boolean) As String
    'on error resume next 'get the 1 byte of a Boolean value
    Stuff_GetBooleanString = String$(1, Chr$(0))
    Call CopyMemory(ByVal Stuff_GetBooleanString, BooleanValue, 1)
End Function

Public Function Stuff_GetStringBoolean(ByVal StringString As String) As Boolean
    'on error resume next
    Call CopyMemory(Stuff_GetStringBoolean, ByVal StringString, 1)
End Function

Public Function Stuff_GetLongArrayString(ByRef LongArray() As Long) As String
    'on error resume next
    Dim ArrayLBound As Long
    Dim ArrayUBound As Long
    Dim Temp As Long
    Dim TempByteArray() As Byte
    'preset
    ArrayLBound = LBound(LongArray())
    ArrayUBound = UBound(LongArray())
    'begin
    ReDim TempByteArray(1 To (ArrayUBound - ArrayLBound + 1&) * 4&) As Byte
    For Temp = ArrayLBound To ArrayUBound
        Call CopyMemory(TempByteArray((Temp - ArrayLBound) * 4& + 1&), LongArray(Temp), 4&)
    Next Temp
    Stuff_GetLongArrayString = String$(UBound(TempByteArray()), Chr$(0))
    Call CopyMemory(ByVal Stuff_GetLongArrayString, TempByteArray(1), Len(Stuff_GetLongArrayString))
End Function

Public Sub Stuff_GetStringLongArray(ByRef LongArray() As Long, ByRef StringString As String, Optional ByVal LongArrayIsFixedFlag As Boolean = False)
    'on error resume next
    Dim Temp As Long
    Dim TempByteArray() As Byte
    'begin
    If LongArrayIsFixedFlag = False Then _
        ReDim LongArray(1 To (-Int(-Len(StringString) / 4&))) As Long
    ReDim TempByteArray(1 To Len(StringString)) As Byte
    Call CopyMemory(TempByteArray(1), ByVal StringString, Len(StringString))
    For Temp = 1& To (-Int(-Len(StringString) / 4&))
        Call CopyMemory(LongArray(Temp), TempByteArray((Temp - 1&) * 4& + 1&), 4&)
    Next Temp
End Sub

'***END OF CONVERSION FUNCTIONS***
'***DRAWING FUNCTIONS***

Public Sub StuffFastLine_SetColor(ByVal hDC As Long, ByVal Color As Long)
    'on error resume next 'if not called before drawing line then line will be black and one pixel width
    Dim StuffFastLinePenHandle As Long
    'begin
    If (StuffFastLinePenHandleUnchanged) Then
        StuffFastLinePenHandle = CreatePen(PS_SOLID, 0&, Color) 'create 1 pixel width pen
        Call DeleteObject(SelectObject(hDC, StuffFastLinePenHandle)) 'previous object was not the original pen
    Else
        StuffFastLinePenHandle = CreatePen(PS_SOLID, 0&, Color) 'create 1 pixel width pen
        StuffFastLinePenHandleUnchanged = SelectObject(hDC, StuffFastLinePenHandle)
    End If
End Sub

Public Sub StuffFastLine_Draw(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    'on error resume next 'around 400 times faster than VB's lame Line method (tested)
    Call MoveToEx(hDC, X1, Y1, 0&)
    Call LineTo(hDC, X2, Y2)
End Sub

Public Sub StuffFastLine_Terminate(ByVal hDC As Long)
    'on error resume next 'to be called when target app is quit
    If (StuffFastLinePenHandleUnchanged) Then
        Call DeleteObject(SelectObject(hDC, StuffFastLinePenHandleUnchanged))
    End If
End Sub

'***END OF DRAWING FUNCTIONS***
'***OTHER FUNCTIONS***

Public Function Stuff_CutNullTermination(ByVal RemoveString As String) As String
    'on error resume next 'call to cut null-terminated strings
    Dim Temp As Long
    'begin
    Temp = InStr(1, RemoveString, Chr$(0), vbBinaryCompare)
    If (Temp) Then
        Stuff_CutNullTermination = Left$(RemoveString, Temp - 1)
    Else
        Stuff_CutNullTermination = RemoveString
    End If
End Function

Public Function STUFF_PROGRAMDIRECTORY() As String
    'on error resume next
    Dim ProgramDirectory As String
    'begin
    ProgramDirectory = App.Path
    If Not (Right$(ProgramDirectory, 1) = "\") Then ProgramDirectory = ProgramDirectory + "\" 'verify
    STUFF_PROGRAMDIRECTORY = ProgramDirectory
End Function

Public Function STUFF_MIN(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error resume next
    If Value1 < Value2 Then
        STUFF_MIN = Value1
    Else
        STUFF_MIN = Value2
    End If
End Function

Public Function STUFF_MAX(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error resume next
    If Value1 > Value2 Then
        STUFF_MAX = Value1
    Else
        STUFF_MAX = Value2
    End If
End Function

Public Function STUFF_DIV(ByVal Value As Long, ByVal Divisor As Long) As Long
    'on error resume next 'how often one number 'goes into' an other
    STUFF_DIV = (Value - (Value Mod Divisor)) \ Divisor
End Function

Public Function STUFF_GETBYTEBITWINDOW(ByVal LongNumber As Long, ByVal FirstBitIndex As Long, ByVal BitNumber As Long) As Byte
    'on error resume next 'take the 8 bits from LongNumber starting at bit offset FirstBitIndex and shifts them down so that they 'fit' into a byte value
    Dim BitFor As Integer
    Dim Temp As Long 'value initially 0
    '
    'NOTE: this function can be used for two purposes:
    '-one possibility to implement the functionality of ReRGB():
    'R = STUFF_GETBYTEBITWINDOW(RGB, 17, 8)
    'G = STUFF_GETBYTEBITWINDOW(RGB, 9, 8)
    'B = STUFF_GETBYTEBITWINDOW(RGB, 1, 8)
    '-create a length string out of a Long-length value:
    'Mid$(HeaderString, 1, 1) = Chr$(STUFF_GETBYTEBITWINDOW(HeaderLength, 25, 8))
    'Mid$(HeaderString, 2, 1) = Chr$(STUFF_GETBYTEBITWINDOW(HeaderLength, 17, 8))
    'Mid$(HeaderString, 3, 1) = Chr$(STUFF_GETBYTEBITWINDOW(HeaderLength, 9, 8))
    'Mid$(HeaderString, 4, 1) = Chr$(STUFF_GETBYTEBITWINDOW(HeaderLength, 1, 8)).
    '
    'verify
    Select Case FirstBitIndex
    Case Is < 1&
        FirstBitIndex = 1&
    Case Is > 25&
        FirstBitIndex = 25&
    End Select
    FirstBitIndex = FirstBitIndex - 1& '1 to 0 based (easier calculation)
    'begin
    For BitFor = 0 To (BitNumber - 1&)
        If (LongNumber And (2& ^ (BitFor + FirstBitIndex))) Then
            Temp = Temp + (2& ^ BitFor) '+ is a little bit faster than Or (tested)
        End If
    Next BitFor
    STUFF_GETBYTEBITWINDOW = CByte(Temp)
End Function

'***END OF OTHER FUNCTIONS***
