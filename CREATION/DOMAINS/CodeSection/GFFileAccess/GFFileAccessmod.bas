Attribute VB_Name = "GFFileAccessmod"
Option Explicit
'(c)2001, 2004 by Louis. Functions for FAST (!) file access.
'GFFileAccess_GetDirFileSizeTotal
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'Get[Total/Avail]DiskSpace (source: drvspace.zip)
Private Declare Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpClusterSectorNumber As Long, lpSectorByteNumber As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'GFFileAccess_GetDirFileSizeTotal
Private Const MAX_PATH = 260
'GFFileAccess_GetDirFileSizeTotal
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
'GFFileAccess_GetDirFileSizeTotal
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
'GFFileAccess_Get[Free/Total]DiskSpace
Private Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Function GFFileAccess_GetDirFileSizeTotal(ByVal DirectoryName As String, ByVal Pattern As String) As Double
    'on error resume next 'returns total file size of all files matching the passed search pattern or -1 for error
    Dim FindFileHandle As Long
    Dim FileSizeTotal As Double
    Dim WIN32_FIND_DATAVar As WIN32_FIND_DATA
    '
    'NOTE: this function is bloody fast, in tests it needed less that 2 seconds
    'on an Athlon 800 to get the total size of 31.461 mp3 files.
    '
    'verify
    If Dir(DirectoryName, vbDirectory) = "" Then
        GFFileAccess_GetDirFileSizeTotal = (-1#) 'error
        Exit Function
    End If
    If Not (Right$(DirectoryName, 1) = "\") Then DirectoryName = DirectoryName + "\"
    'begin
    FindFileHandle = FindFirstFile(DirectoryName + Pattern, WIN32_FIND_DATAVar)
    If FindFileHandle > 0& Then
ReDo:
        FileSizeTotal = FileSizeTotal + CDbl(WIN32_FIND_DATAVar.nFileSizeLow)
        If FindNextFile(FindFileHandle, WIN32_FIND_DATAVar) > 0& Then GoTo ReDo:
        Call FindClose(FindFileHandle)
        GFFileAccess_GetDirFileSizeTotal = FileSizeTotal 'ok
        Exit Function
    Else
        GFFileAccess_GetDirFileSizeTotal = (-1#) 'error
        Exit Function
    End If
End Function

'NOTE: GFFileAccess_IsFileExisting() requires Attributes to be passed, DirSave() doesn't.

Public Function GFFileAccess_IsFileExisting(ByVal DirectoryName As String, ByVal Pattern As String) As Boolean
    'on error resume next 'returns True if directory contains files that match the pattern, False if not
    Dim FindFileHandle As Long
    Dim WIN32_FIND_DATAVar As WIN32_FIND_DATA
    '
    'NOTE: use this function if it must be determinated if a large amount
    'of directories contain files of a special type (pattern).
    'Directory must exist and must be back-slash terminated.
    '
    'begin
    FindFileHandle = FindFirstFile(DirectoryName + Pattern, WIN32_FIND_DATAVar)
    GFFileAccess_IsFileExisting = (FindFileHandle > 0&)
    Call FindClose(FindFileHandle)
End Function

Public Function GFFileAccess_DirSave(ByVal PathName As String, ByVal Attributes As Integer) As String
    On Error GoTo Error: 'important
    '
    'NOTE: Dir() raises an error if PathName represents a cdrom drive
    'and the cd is not inserted (damn VB!). Use this function rather than Dir().
    '
    GFFileAccess_DirSave = Dir(PathName, Attributes) 'ok
    Exit Function
Error:
    GFFileAccess_DirSave = "" 'error
    Exit Function
End Function

Public Function DirSave(ByVal PathName As String, Optional ByVal Attributes As Integer = vbNormal) As String
    On Error GoTo Error: 'important
    '
    'NOTE: Dir() raises an error if PathName represents a cdrom drive
    'and the cd is not inserted (damn VB!). Use this function rather than Dir().
    '
    DirSave = Dir(PathName, Attributes) 'ok
    Exit Function
Error:
    DirSave = "" 'error
    Exit Function
End Function

Public Function GetAttrSave(ByRef PathName As String) As VbFileAttribute
    On Error GoTo Error: 'important
    '
    'NOTE: GetAttr() raises an error if PathName is a cdrom drive
    'with no cd inserted. Use GetAttrSave() instead of GetAttr() if
    'PathName could be a cdrom drive.
    '
    GetAttrSave = GetAttr(PathName)
    Exit Function
Error:
    GetAttrSave = vbNormal
    Exit Function
End Function
    
'***DISK SPACE FUNCTIONS***
'NOTE: the following two functions are to be used to dterminate the free or total space available on a special drive.
'The two functions were created out of the Noname99 functions GetAvailableDiskSpace() and GetTotalDiskSpace().
'The two functions use the API functions GetDiskFreeSpace() and GetDiskFreeSpaceEx().
'The largest detrminable size using GetDiskFreeSpace() is 2 GB, from Win95 OSR 2 on GetDiskFreeSpaceEx()
'can be used to retreive sizes above 2 GB.
'Code was partially taken from http://www.vbapi.com/ref/g/getdiskfreespaceex.html (03.02.2002).

Public Function GFFileAccess_GetFreeDiskSpace(ByVal DiskName As String) As Double
    On Error GoTo Error: 'important; returns free disk space in bytes
    Dim BytesFreeToUser As ULARGE_INTEGER
    Dim BytesTotal As ULARGE_INTEGER
    Dim BytesFree As ULARGE_INTEGER
    Dim TempCurrency As Currency
    Dim Temp As Long
    'begin
    Call GetDiskFreeSpaceEx(DiskName, BytesFreeToUser, BytesTotal, BytesFree)
    Call CopyMemory(TempCurrency, BytesFreeToUser, 8) 'taken from http://www.vbapi.com/ref/g/getdiskfreespaceex.html
    GFFileAccess_GetFreeDiskSpace = CDbl(TempCurrency * 10000@)
    Exit Function
Error: 'on Win95 OSR 1
    Dim ClusterSectorNumber As Long
    Dim SectorByteNumber As Long
    Dim ClusterNumberFree As Long
    Dim ClusterNumberTotal As Long
    Call GetDiskFreeSpace(DiskName, ClusterSectorNumber, SectorByteNumber, ClusterNumberFree, ClusterNumberTotal)
    GFFileAccess_GetFreeDiskSpace = CDbl(ClusterSectorNumber) * CDbl(SectorByteNumber) * CDbl(ClusterNumberFree)
    Exit Function
End Function

Public Function GFFileAccess_GetTotalDiskSpace(ByVal DiskName As String) As Double
    On Error GoTo Error: 'important; returns total disk space in bytes
    Dim BytesFreeToUser As ULARGE_INTEGER
    Dim BytesTotal As ULARGE_INTEGER
    Dim BytesFree As ULARGE_INTEGER
    Dim TempCurrency As Currency
    Dim Temp As Long
    'begin
    Call GetDiskFreeSpaceEx(DiskName, BytesFreeToUser, BytesTotal, BytesFree)
    Call CopyMemory(TempCurrency, BytesTotal, 8) 'taken from http://www.vbapi.com/ref/g/getdiskfreespaceex.html
    GFFileAccess_GetTotalDiskSpace = CDbl(TempCurrency * 10000@)
    Exit Function
Error: 'on Win95 OSR 1
    Dim ClusterSectorNumber As Long
    Dim SectorByteNumber As Long
    Dim ClusterNumberFree As Long
    Dim ClusterNumberTotal As Long
    Call GetDiskFreeSpace(DiskName, ClusterSectorNumber, SectorByteNumber, ClusterNumberFree, ClusterNumberTotal)
    GFFileAccess_GetTotalDiskSpace = CDbl(ClusterSectorNumber) * CDbl(SectorByteNumber) * CDbl(ClusterNumberTotal)
    Exit Function
End Function

