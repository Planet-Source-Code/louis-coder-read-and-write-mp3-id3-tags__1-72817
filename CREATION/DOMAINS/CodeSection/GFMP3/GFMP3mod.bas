Attribute VB_Name = "GFMP3mod"
Option Explicit
'(c)2002, 2003, 2004 by Louis.
'
'NOTE: This module contains procedures to read and write mp3 tags.
'
'GFShrinkFile
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'general use
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'GFShrinkFile
Private Const OFS_MAXPATHNAME = 128
Private Const OF_READWRITE = &H2
Private Const FILE_BEGIN = 0
'GFShrinkFile
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
'ShiftBits (copied)
'Enum for the Bit Shifting Direction
Private Enum eBitShiftDir
    eShiftLeft = 1
    eShiftRight = 2
End Enum

'***************************************MP3 TAG 1**************************************
'NOTE: the MP3 TAG subs are used to read/write the TAG (v1.0) of an MP3 file.
'NOTE: the MP3 TAG functions were copied from MP3 Namer 1 (04-27-2001).
'NOTE: there are two versions of the MP3 TAG functions: string and byte functions.
'MP3 Namer uses the byte functions, the other ones are included for future usage.
'Note that MP3TAG1_GetMP3TAGNewStartPos() is used by both string and byte functions.

Public Sub MP3TAG1_Read(ByVal MP3Name As String, ByRef MP3SongName As String, ByRef MP3ArtistName As String, ByRef MP3AlbumName As String, ByRef MP3YearName As String, ByRef MP3Comment As String, ByRef MP3Genre As Byte)
    'on error Resume Next 'reads the MP3 TAG if existing
    Dim MP3NameFileNumber As Integer
    Dim MP3NameString As String
    Dim MP3TAGStringStartPos As Long
    'preset
    MP3NameFileNumber = FreeFile(0)
    MP3SongName = ""
    MP3ArtistName = ""
    MP3AlbumName = ""
    MP3YearName = ""
    MP3Comment = ""
    MP3Genre = 0
    'begin
    If Not ((DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "")) Then 'verify
        Open MP3Name For Binary As #MP3NameFileNumber
            Select Case LOF(MP3NameFileNumber)
            Case 0 'nothing to read
                Close #MP3NameFileNumber 'important
                Exit Sub
            Case Is < 128
                MP3NameString = String$(LOF(MP3NameFileNumber), Chr$(0))
            Case Else
                MP3NameString = String$(128, Chr$(0))
            End Select
            Get #MP3NameFileNumber, LOF(MP3NameFileNumber) - Len(MP3NameString) + 1, MP3NameString
        Close #MP3NameFileNumber
        MP3TAGStringStartPos = InStr(1, MP3NameString, "TAG", vbTextCompare)
        If Not (MP3TAGStringStartPos = 0) Then 'check if TAG was found
            'title
            MP3TAGStringStartPos = MP3TAGStringStartPos + Len("TAG")
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3SongName = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'artist
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3ArtistName = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'album
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3AlbumName = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'year
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3YearName = Mid$(MP3NameString, MP3TAGStringStartPos, 4) 'string verified later
            'comment
            MP3TAGStringStartPos = MP3TAGStringStartPos + 4
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3Comment = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'end of string
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3Genre = Asc(Mid$(MP3NameString, MP3TAGStringStartPos, 1))
        Else
            GoTo Leave:
        End If
    Else
        MsgBox "internal error in MP3TAG1_Read(): file " + Left$(MP3Name, 1024) + " not found !", vbOKOnly + vbExclamation
    End If
Leave:
    Call MP3TAG1_Read_VerifyMP3TAGItem(MP3SongName)
    Call MP3TAG1_Read_VerifyMP3TAGItem(MP3ArtistName)
    Call MP3TAG1_Read_VerifyMP3TAGItem(MP3AlbumName)
    Call MP3TAG1_Read_VerifyMP3TAGItem(MP3YearName)
    Call MP3TAG1_Read_VerifyMP3TAGItem(MP3Comment)
    Exit Sub
End Sub

Public Function MP3TAG1_ReadByte(ByVal MP3Name As String, ByRef MP3SongName() As Byte, ByRef MP3ArtistName() As Byte, ByRef MP3AlbumName() As Byte, ByRef MP3YearName() As Byte, ByRef MP3Comment() As Byte, ByRef MP3Genre As Byte, ByVal TAGSTRINGLENGTH As Long) As Boolean
    On Error GoTo Error: 'reads the MP3 TAG if existing; returns True if TAG was partially read, False if not
    Dim MP3NameString As String
    Dim MP3NameFileNumber As Integer
    Dim MP3TAGStringStartPos As Long
    Dim Tempstr$
    'preset
    'MP3SongName = ""
    'MP3ArtistName = ""
    'MP3AlbumName = ""
    'MP3YearName = ""
    'MP3Comment = ""
    'MP3Genre = 0
    MP3NameFileNumber = FreeFile(0)
    'begin
    If Not ((DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "")) Then 'verify
        Open MP3Name For Binary As #MP3NameFileNumber
            Select Case LOF(MP3NameFileNumber)
            Case 0 'nothing to read
                Close #MP3NameFileNumber
                GoTo Error:
            Case Is < 128
                MP3NameString = String$(LOF(1), Chr$(0))
            Case Else
                MP3NameString = String$(128, Chr$(0))
            End Select
            Get #MP3NameFileNumber, LOF(MP3NameFileNumber) - Len(MP3NameString) + 1, MP3NameString
        Close #MP3NameFileNumber
        MP3TAGStringStartPos = InStr(1, MP3NameString, "TAG", vbTextCompare)
        If Not (MP3TAGStringStartPos = 0) Then 'check if TAG was found
            'song
            MP3TAGStringStartPos = MP3TAGStringStartPos + Len("TAG")
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            Tempstr$ = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'Call RemoveChr0(Tempstr$) 'important (or ByteString() functions will fail)
            Call CopyMemory(MP3SongName(1), ByVal Tempstr$, MIN(TAGSTRINGLENGTH, 30))
            Call RemoveChr0Byte(MP3SongName(), MIN(TAGSTRINGLENGTH, 30)) 'same as the string function would do
            'artist
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            Tempstr$ = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'Call RemoveChr0(Tempstr$) 'important (or ByteString() functions will fail)
            Call CopyMemory(MP3ArtistName(1), ByVal Tempstr$, MIN(TAGSTRINGLENGTH, 30))
            Call RemoveChr0Byte(MP3ArtistName(), MIN(TAGSTRINGLENGTH, 30)) 'same as the string function would do
            'album
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            Tempstr$ = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'Call RemoveChr0(Tempstr$) 'important (or ByteString() functions will fail)
            Call CopyMemory(MP3AlbumName(1), ByVal Tempstr$, MIN(TAGSTRINGLENGTH, 30))
            Call RemoveChr0Byte(MP3AlbumName(), MIN(TAGSTRINGLENGTH, 30)) 'same as the string function would do
            'year
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            Tempstr$ = Mid$(MP3NameString, MP3TAGStringStartPos, 4) 'string verified later
            Call RemoveChr0(Tempstr$) 'important (or ByteString() functions will fail)
            Call CopyMemory(MP3YearName(1), ByVal Tempstr$, MIN(TAGSTRINGLENGTH, 4))
            Call RemoveChr0Byte(MP3YearName(), MIN(TAGSTRINGLENGTH, 4)) 'same as the string function would do
            'comment
            MP3TAGStringStartPos = MP3TAGStringStartPos + 4
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            Tempstr$ = Mid$(MP3NameString, MP3TAGStringStartPos, 30) 'string verified later
            'Call RemoveChr0(Tempstr$) 'important (or ByteString() functions will fail)
            Call CopyMemory(MP3Comment(1), ByVal Tempstr$, MIN(TAGSTRINGLENGTH, 30))
            Call RemoveChr0Byte(MP3Comment(), MIN(TAGSTRINGLENGTH, 30)) 'same as the string function would do
            'genre
            MP3TAGStringStartPos = MP3TAGStringStartPos + 30
            If MP3TAGStringStartPos > Len(MP3NameString) Then GoTo Leave:
            MP3Genre = Asc(Mid$(MP3NameString, MP3TAGStringStartPos, 1))
        Else
            GoTo Error:
        End If
    Else
        MsgBox "internal error in MP3TAG1_ReadByte(): file " + Left$(MP3Name, 1024) + " not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
Leave:
    Call MP3TAG1_Read_VerifyMP3TAGItemByte(MP3SongName(), TAGSTRINGLENGTH)
    Call MP3TAG1_Read_VerifyMP3TAGItemByte(MP3ArtistName(), TAGSTRINGLENGTH)
    Call MP3TAG1_Read_VerifyMP3TAGItemByte(MP3AlbumName(), TAGSTRINGLENGTH)
    Call MP3TAG1_Read_VerifyMP3TAGItemByte(MP3YearName(), TAGSTRINGLENGTH)
    Call MP3TAG1_Read_VerifyMP3TAGItemByte(MP3Comment(), TAGSTRINGLENGTH)
    MP3TAG1_ReadByte = True 'ok
    Exit Function
Error: 'error or no TAG found
    Close #MP3NameFileNumber 'make sure file is closed
    MP3TAG1_ReadByte = False 'error
    Exit Function
End Function

Public Sub MP3TAG1_Read_VerifyMP3TAGItem(ByRef MP3TAGItem As String)
    'on error Resume Next 'cuts invalid chars at start/end of passed string
    Dim MP3TAGTemp As Long
    Dim MP3TAGItemNewStartPos As Long
    Dim MP3TAGItemNewEndPos As Long
    'verify
    If MP3TAGItem = "" Then Exit Sub 'Mid$() would fail
    'preset
    MP3TAGItemNewStartPos = 1
    MP3TAGItemNewEndPos = Len(MP3TAGItem)
    'begin; calculate
    For MP3TAGTemp = 1 To Len(MP3TAGItem)
        Select Case Asc(Mid$(MP3TAGItem, MP3TAGTemp, 1))
        Case Is <= 32
           MP3TAGItemNewStartPos = MP3TAGTemp + 1
        Case Else
            Exit For
        End Select
    Next MP3TAGTemp
    'cut
    If Not (MP3TAGItemNewStartPos > Len(MP3TAGItem)) Then
        MP3TAGItem = Right$(MP3TAGItem, Len(MP3TAGItem) - MP3TAGItemNewStartPos + 1)
    Else
        MP3TAGItem = "" 'reset (error)
    End If
    'calculate
    For MP3TAGTemp = Len(MP3TAGItem) To 1 Step (-1)
        Select Case Asc(Mid$(MP3TAGItem, MP3TAGTemp, 1))
        Case Is <= 32
           MP3TAGItemNewEndPos = MP3TAGTemp - 1
        Case Else
            Exit For
        End Select
    Next MP3TAGTemp
    'cut
    If Not (MP3TAGItemNewEndPos < 1) Then
        MP3TAGItem = Left$(MP3TAGItem, MP3TAGItemNewEndPos)
    Else
        MP3TAGItem = "" 'reset (error)
    End If
End Sub

Public Sub MP3TAG1_Read_VerifyMP3TAGItemByte(ByRef MP3TAGItem() As Byte, ByVal TAGSTRINGLENGTH As Long)
    'on error Resume Next 'cuts invalid chars at start/end of passed string
    Dim MP3TAGTemp As Long
    Dim MP3TAGItemLength As Long
    Dim MP3TAGItemNewStartPos As Long
    Dim MP3TAGItemNewEndPos As Long
    'verify
    If TAGSTRINGLENGTH = 0 Then Exit Sub 'Mid$() would fail
    'preset
    MP3TAGItemLength = GETBYTESTRINGLENGTH(MP3TAGItem())
    MP3TAGItemNewStartPos = 1
    MP3TAGItemNewEndPos = MP3TAGItemLength
    'begin, calculate
    For MP3TAGTemp = 1 To MP3TAGItemLength
        Select Case MP3TAGItem(MP3TAGTemp)
        Case Is <= 32
           MP3TAGItemNewStartPos = MP3TAGTemp + 1
        Case Else
            Exit For
        End Select
    Next MP3TAGTemp
    'cut
    If Not (MP3TAGItemNewStartPos > MP3TAGItemLength) Then
        Call CopyMemory(MP3TAGItem(1), MP3TAGItem(MP3TAGItemNewStartPos), MP3TAGItemLength - MP3TAGItemNewStartPos + 1)
        Call BYTESTRINGLEFT(MP3TAGItem(), MP3TAGItemLength - MP3TAGItemNewStartPos + 1)
    Else
        Call BYTESTRINGLEFT(MP3TAGItem(), 0) 'error
    End If
    'MP3TAGItemLength = GETBYTESTRINGLENGTH(MP3TAGItem()) 'refresh 'no, not used any more
    'calculate
    For MP3TAGTemp = MP3TAGItemLength To 1 Step (-1)
        Select Case MP3TAGItem(MP3TAGTemp)
        Case Is <= 32
           MP3TAGItemNewEndPos = MP3TAGTemp - 1
        Case Else
            Exit For
        End Select
    Next MP3TAGTemp
    'cut
    If Not (MP3TAGItemNewEndPos < 1) Then
        'no data must be moved
        Call BYTESTRINGLEFT(MP3TAGItem(), MP3TAGItemNewEndPos)
    Else
        Call BYTESTRINGLEFT(MP3TAGItem(), 0) 'error
    End If
End Sub

Public Function MP3TAG1_Write(ByVal MP3Name As String, ByVal MP3SongName As String, ByVal MP3ArtistName As String, ByVal MP3AlbumName As String, ByVal MP3YearName As String, ByVal MP3Comment As String, ByVal MP3Genre As Byte) As Boolean
    On Error GoTo Error: 'important; writes an ID3v1 TAG over existing TAG/at file end; returns True for success or False for error
    Dim MP3TAGStartPos As Long
    Dim MP3TAGString As String
    Dim MP3NameFileNumber As Integer
    'preset
    MP3NameFileNumber = FreeFile(0)
    'begin
    If Not ((DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "")) Then 'verify
        'remove existing MP3 TAG
        MP3TAGStartPos = MP3TAG1_GetMP3TAGNewStartPos(MP3Name)
        Select Case MP3TAGStartPos 'verify
        Case 0
            Exit Function 'error
        Case Is < FileLen(MP3Name)
            If GFShrinkFile(MP3Name, (MP3TAGStartPos - 1)) = True Then
                'ok
            Else
                Exit Function 'error
            End If
        Case Else
            'ok (do nothing)
        End Select
        'append MP3 TAG
        Open MP3Name For Append As #MP3NameFileNumber 'if an error during opening file occurs then jump to Error:
            MP3TAGString = MP3TAGString + "TAG"
            MP3TAGString = MP3TAGString + Left$(MP3SongName, 30) + String$(30 - Len(Left$(MP3SongName, 30)), Chr$(0))
            MP3TAGString = MP3TAGString + Left$(MP3ArtistName, 30) + String$(30 - Len(Left$(MP3ArtistName, 30)), Chr$(0))
            MP3TAGString = MP3TAGString + Left$(MP3AlbumName, 30) + String$(30 - Len(Left$(MP3AlbumName, 30)), Chr$(0))
            MP3TAGString = MP3TAGString + Left$(MP3YearName, 4) + String$(4 - Len(Left$(MP3YearName, 4)), Chr$(0))
            MP3TAGString = MP3TAGString + Left$(MP3Comment, 30) + String$(30 - Len(Left$(MP3Comment, 30)), Chr$(0))
            MP3TAGString = MP3TAGString + Chr$(MP3Genre)
            Print #MP3NameFileNumber, MP3TAGString;
        Close #MP3NameFileNumber
    Else
        GoTo Error:
    End If
    MP3TAG1_Write = True 'ok
    Exit Function
Error:
    Close #MP3NameFileNumber 'make sure file is closed
    MP3TAG1_Write = False 'error
    Exit Function
End Function

Public Function MP3TAG1_GetMP3TAGNewStartPos(ByVal MP3Name As String) As Long
    'on error Resume Next 'returns the position the new TAG must be begun at or 0 for error
    Dim MP3NameString As String
    Dim MP3NameSize As Long
    Dim MP3TAGStringStartPos As Long
    Dim MP3NameFileNumber As Integer
    'preset
    MP3NameFileNumber = FreeFile(0)
    'verify
    If ((DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "")) Then 'verify
        MP3TAG1_GetMP3TAGNewStartPos = 0 'error
        Exit Function
    End If
    'begin
    MP3NameSize = FileLen(MP3Name)
    Open MP3Name For Binary As #MP3NameFileNumber
        Select Case LOF(MP3NameFileNumber)
        Case 0 'nothing to read
            MP3TAG1_GetMP3TAGNewStartPos = MP3NameSize 'ok
            Close #MP3NameFileNumber
            Exit Function
        Case Is < 128
            MP3NameString = String$(LOF(MP3NameFileNumber), Chr$(0))
        Case Else
            MP3NameString = String$(128, Chr$(0))
        End Select
        Get #MP3NameFileNumber, LOF(MP3NameFileNumber) - Len(MP3NameString) + 1, MP3NameString
    Close #MP3NameFileNumber
    MP3TAGStringStartPos = InStr(1, MP3NameString, "TAG", vbTextCompare)
    If Not (MP3TAGStringStartPos = 0) Then 'check if TAG was found
        MP3TAG1_GetMP3TAGNewStartPos = (MP3NameSize - Len(MP3NameString) + 1) + (MP3TAGStringStartPos - 1) 'ok
        Exit Function
    Else
        MP3TAG1_GetMP3TAGNewStartPos = MP3NameSize 'ok
        Exit Function
    End If
    Exit Function
End Function

'***********************************END OF MP3 TAG 1***********************************
'***************************************MP3 TAG 2**************************************
'NOTE: use these functions to read a version 2 TAG (implemented 12.01.2003).
'
'IMPORTANT: TAG frames are only read and written correctly when they aren't
'longer than approx. 256 ^ 3 bytes (so that the highest-order byte of the size isn't
'used).

Public Function MP3TAG2_ReadByte(ByVal MP3Name As String, ByRef MP3SongName() As Byte, ByRef MP3ArtistName() As Byte, ByRef MP3AlbumName() As Byte, ByRef MP3YearName() As Byte, ByRef MP3Comment() As Byte, ByRef MP3GenreName() As Byte, ByRef MP3Genre As Byte, ByVal TAGSTRINGLENGTH As Long) As Boolean
    On Error GoTo Error: 'important (if a file is write-protected or damaged); reads song, artist, album, year name, comment and genre name and -byte of an ID3v2(.2/.3) TAG; returns True for success or False for error
    Dim MP3NameFileNumber As Integer
    Dim HeaderLength As Double 'use double to avoid overflow error
    Dim HeaderMajorVersion As Byte
    Dim HeaderString As String
    Dim HeaderByteString() As Byte
    Dim FrameDataLength As Long
    Dim FrameDataLengthEndPos As Long
    Dim FrameStartPos As Long
    Dim TAGSongFrameName As String
    Dim TAGArtistFrameName As String
    Dim TAGAlbumFrameName As String
    Dim TAGYearFrameName As String
    Dim TAGCommentFrameName As String
    Dim TAGGenreFrameName As String
    Dim TAGFramesStartPos As Long
    Dim EndPos As Long
    Dim Tempstr$
    Dim TempByte As Byte
    Dim TempByteString() As Byte
    '
    'NOTE: this code was copied from a sample downloaded from www.pscode.com (12.01.2003).
    'For further information read the HTML page in the GFMP3 directory.
    'NOTE: in the ID3v2 TAG there's a (self-called) genre-name and a genre-byte.
    'The var (byte string) saving genre-name is called MP3GenreName() and the var saving
    'the genre-byte is just called MP3Genre.
    'NOTE: the calling procedure must reset the passed byte string, not done by this function.
    '
    'preset
    MP3NameFileNumber = FreeFile(0)
    'begin
    If Not ((DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "")) Then 'verify
        'read header string
        Open MP3Name For Binary As #MP3NameFileNumber
            'verify file length
            If LOF(MP3NameFileNumber) < 32 Then GoTo Error: 'there must be something in it to read first (11) header bytes
            'read header 'info block'
            ReDim TempByteString(1 To 10) As Byte
            Get #MP3NameFileNumber, 1, TempByteString()
            'check for string 'ID3'
            If Not (TempByteString(1) = 73) Then GoTo Error: 'no 'I'
            If Not (TempByteString(2) = 68) Then GoTo Error: 'no 'D'
            If Not (TempByteString(3) = 51) Then GoTo Error 'no '3'
            'read header version
            HeaderMajorVersion = TempByteString(4)
            'byte 5: revision number
            'byte 6: flags
            'read header size
            HeaderLength = 0# 'reset
            HeaderLength = HeaderLength + (CDbl(TempByteString(7)) * 20917152#)
            HeaderLength = HeaderLength + (CDbl(TempByteString(8)) * 16384#)
            HeaderLength = HeaderLength + (CDbl(TempByteString(9)) * 128#)
            HeaderLength = HeaderLength + (CDbl(TempByteString(10)))
            If (HeaderLength < 1#) Or (HeaderLength > LOF(MP3NameFileNumber)) Or (HeaderLength > 2147483647#) Then 'verify
                GoTo Error:
            End If
            'read header string
            HeaderString = String$(CLng(HeaderLength), Chr$(0))
            Get #MP3NameFileNumber, 11, HeaderString
            Call GETBYTESTRINGFROMSTRING(Len(HeaderString), HeaderByteString(), HeaderString)
            '
        Close #MP3NameFileNumber
        'allocate header string
        '
        'NOTE: the blocks that store one tag item's data will be called the 'Frames'.
        'Every Frame is identified by its 'Frame name'.
        '
        Select Case HeaderMajorVersion 'two versions are supported
        Case 2 'ID3v2.2
            TAGSongFrameName = "TT2"
            TAGArtistFrameName = "TOA"
            TAGAlbumFrameName = "TAL"
            TAGYearFrameName = "TYE"
            TAGCommentFrameName = "COM"
            TAGGenreFrameName = "TCO"
            TAGFramesStartPos = 7
            FrameDataLengthEndPos = 5 'last char of the frame data length
        Case 3 'ID3v2.3
            TAGSongFrameName = "TIT2"
            TAGArtistFrameName = "TPE1"
            TAGAlbumFrameName = "TALB"
            TAGYearFrameName = "TYER"
            TAGCommentFrameName = "COMM"
            TAGGenreFrameName = "TCON"
            TAGFramesStartPos = 11
            FrameDataLengthEndPos = 7 'last char of the frame data length
        Case Else
            GoTo Error: 'header version is not supported
        End Select
        'read song, artist, album, year name, comment and genre
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGSongFrameName)
        If (FrameStartPos) Then 'verify
            'read the title
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one (all tested with Winamp3)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadArtistName:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3SongName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadArtistName:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGArtistFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadAlbumName:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3ArtistName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadAlbumName:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGAlbumFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadYearName:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3AlbumName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadYearName:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGYearFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadComment:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3YearName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadComment:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGCommentFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadGenre:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3Comment(1), HeaderByteString(FrameStartPos + TAGFramesStartPos + 4), MIN(FrameDataLength - 4, TAGSTRINGLENGTH)) 'add and subtract 4 (tested)
        End If
ReadGenre:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGGenreFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo Leave:
            End Select
            Tempstr$ = Mid$(HeaderString, FrameStartPos + TAGFramesStartPos, FrameDataLength)
            If HeaderByteString(FrameStartPos + TAGFramesStartPos) = 40 Then '(
                EndPos = InStr(1, Tempstr$, ")", vbBinaryCompare)
                If (EndPos) Then
                    '
                    'NOTE: there's an ID3v2 genre that has the following format:
                    '(0) Blues
                    '(12) Rock
                    '(125) Dance Hall.
                    '
                    MP3Genre = CByte(MIN(MAX(Val(Mid$(Tempstr$, 2, EndPos - 2 + 1)), 0), 255))
                    Tempstr$ = Trim$(Mid$(Tempstr$, EndPos + 1))
                    Call GETFIXEDBYTESTRINGFROMSTRING(TAGSTRINGLENGTH, MP3GenreName(), Tempstr$)
                Else
                    MP3Genre = 0 'reset (unknown)
                    Call GETFIXEDBYTESTRINGFROMSTRING(TAGSTRINGLENGTH, MP3GenreName(), Tempstr$)
                End If
            Else
                MP3Genre = 0 'reset (unknown)
                Call GETFIXEDBYTESTRINGFROMSTRING(TAGSTRINGLENGTH, MP3GenreName(), Tempstr$)
            End If
        End If
Leave:
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3SongName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3ArtistName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3AlbumName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3YearName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3Comment(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3GenreName(), TAGSTRINGLENGTH)
    Else
        MsgBox "internal error in MP3TAG2_ReadByte(): file " + Left$(MP3Name, 1024) + " not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    MP3TAG2_ReadByte = True 'ok
    Exit Function
Error: 'error or no TAG found
    Close #MP3NameFileNumber 'make sure file is closed
    MP3TAG2_ReadByte = False 'error
    Exit Function
End Function

Public Function MP3TAG2_WriteByte(ByVal MP3Name As String, ByRef MP3SongName() As Byte, ByRef MP3ArtistName() As Byte, ByRef MP3AlbumName() As Byte, ByRef MP3YearName() As Byte, ByRef MP3Comment() As Byte, ByRef MP3GenreName() As Byte, ByRef MP3Genre As Byte) As Boolean
    'on error resume next
    MP3TAG2_WriteByte = False '[not supported yet]
End Function

Public Sub MP3TAG2_Read_VerifyMP3TAGItemByte(ByRef MP3TAGItem() As Byte, ByVal TAGSTRINGLENGTH As Long)
    'on error resume next 'verifies there's no Chr$(0) in passed item (except 'byte string fill-Chr$(0)s')
    Call MP3TAG1_Read_VerifyMP3TAGItemByte(MP3TAGItem(), TAGSTRINGLENGTH)
End Sub

Public Function MP3TAG2_ReadByteEx(ByVal MP3Name As String, ByRef MP3SongName() As Byte, ByRef MP3ArtistName() As Byte, ByRef MP3AlbumName() As Byte, ByRef MP3YearName() As Byte, ByRef MP3Comment() As Byte, ByRef MP3GenreName() As Byte, ByRef MP3Genre As Byte, _
    ByRef MP3Composer() As Byte, ByRef MP3OriginalArtist() As Byte, ByRef MP3Copyright() As Byte, ByRef MP3URL() As Byte, ByRef MP3EncodedBy() As Byte, ByVal TAGSTRINGLENGTH As Long) As Boolean
    On Error GoTo Error: 'important (if a file is write-protected or damaged); reads song, artist, album, year name, comment and genre name and -byte of an ID3v2(.2/.3) TAG; returns True for success or False for error
    Dim MP3NameFileNumber As Integer
    Dim HeaderLength As Double 'use double to avoid overflow error
    Dim HeaderMajorVersion As Byte
    Dim HeaderString As String
    Dim HeaderByteString() As Byte
    Dim FrameDataLength As Long 'length of frame data
    Dim FrameDataLengthEndPos As Long 'points to last char of frame size identifier
    Dim FrameStartPos As Long
    Dim TAGSongFrameName As String
    Dim TAGArtistFrameName As String
    Dim TAGAlbumFrameName As String
    Dim TAGYearFrameName As String
    Dim TAGCommentFrameName As String
    Dim TAGGenreFrameName As String
    Dim TAGComposerFrameName As String
    Dim TAGOriginalArtistFrameName As String
    Dim TAGCopyrightFrameName As String
    Dim TAGURLFrameName As String
    Dim TAGEncodedByFrameName As String
    Dim TAGFramesStartPos As Long
    Dim EndPos As Long
    Dim Tempstr$
    Dim TempByte As Byte
    Dim TempByteString() As Byte
    '
    'NOTE: this code was copied from a sample downloaded from www.pscode.com (12.01.2003).
    'For further information read the HTML page in the GFMP3 directory.
    'NOTE: in the ID3v2 TAG there's a (self-called) genre-name and a genre-byte.
    'The var (byte string) saving genre-name is called MP3GenreName() and the var saving
    'the genre-byte is just called MP3Genre.
    'NOTE: the calling procedure must reset the passed byte string, not done by this function.
    '
    'preset
    MP3NameFileNumber = FreeFile(0)
    'begin
    If Not ((DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "")) Then 'verify
        'read header string
        Open MP3Name For Binary As #MP3NameFileNumber
            'verify file length
            If LOF(MP3NameFileNumber) < 32 Then GoTo Error: 'there must be something in it to read first (11) header bytes
            'read header 'info block'
            ReDim TempByteString(1 To 10) As Byte
            Get #MP3NameFileNumber, 1, TempByteString()
            'check for string 'ID3'
            If Not (TempByteString(1) = 73) Then GoTo Error: 'no 'I'
            If Not (TempByteString(2) = 68) Then GoTo Error: 'no 'D'
            If Not (TempByteString(3) = 51) Then GoTo Error 'no '3'
            'read header version
            HeaderMajorVersion = TempByteString(4)
            'byte 5: revision number
            'byte 6: flags
            'read header size
            HeaderLength = 0# 'reset
            HeaderLength = HeaderLength + (CDbl(TempByteString(7)) * 20917152#)
            HeaderLength = HeaderLength + (CDbl(TempByteString(8)) * 16384#)
            HeaderLength = HeaderLength + (CDbl(TempByteString(9)) * 128#)
            HeaderLength = HeaderLength + (CDbl(TempByteString(10)))
            If (HeaderLength < 1#) Or (HeaderLength > LOF(MP3NameFileNumber)) Or (HeaderLength > 2147483647#) Then 'verify
                GoTo Error:
            End If
            'read header string
            HeaderString = String$(CLng(HeaderLength), Chr$(0))
            Get #MP3NameFileNumber, 11, HeaderString
            Call GETBYTESTRINGFROMSTRING(Len(HeaderString), HeaderByteString(), HeaderString)
            '
        Close #MP3NameFileNumber
        'allocate header string
        '
        'NOTE: the blocks that store one tag item's data will be called the 'Frames'.
        'Every Frame is identified by its 'Frame name'.
        '
        Select Case HeaderMajorVersion 'two versions are supported
        Case 2 'ID3v2.2
            TAGSongFrameName = "TT2"
            TAGArtistFrameName = "TOA"
            TAGAlbumFrameName = "TAL"
            TAGYearFrameName = "TYE"
            TAGCommentFrameName = "COM"
            TAGGenreFrameName = "TCO"
            TAGComposerFrameName = Chr$(255) + Chr$(0) + Chr$(255) + Chr$(0) 'not supported, use any string that should not be found
            TAGOriginalArtistFrameName = Chr$(255) + Chr$(0) + Chr$(255) + Chr$(0) 'not supported, use any string that should not be found
            TAGCopyrightFrameName = Chr$(255) + Chr$(0) + Chr$(255) + Chr$(0) 'not supported, use any string that should not be found
            TAGURLFrameName = Chr$(255) + Chr$(0) + Chr$(255) + Chr$(0) 'not supported, use any string that should not be found
            TAGEncodedByFrameName = Chr$(255) + Chr$(0) + Chr$(255) + Chr$(0) 'not supported, use any string that should not be found
            TAGFramesStartPos = 7
            FrameDataLengthEndPos = 5 'last char of the frame data length
        Case 3 'ID3v2.3
            TAGSongFrameName = "TIT2"
            TAGArtistFrameName = "TPE1"
            TAGAlbumFrameName = "TALB"
            TAGYearFrameName = "TYER"
            TAGCommentFrameName = "COMM"
            TAGGenreFrameName = "TCON"
            TAGComposerFrameName = "TCOM"
            TAGOriginalArtistFrameName = "TOPE"
            TAGCopyrightFrameName = "TCOP"
            TAGURLFrameName = "WXXX"
            TAGEncodedByFrameName = "TENC"
            TAGFramesStartPos = 11
            FrameDataLengthEndPos = 7 'last char of the frame data length
        Case Else
            GoTo Error: 'header version is not supported
        End Select
        'read song, artist, album, year name, comment and genre
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGSongFrameName)
        If (FrameStartPos) Then 'verify
            'read the title
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadArtistName:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3SongName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadArtistName:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGArtistFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadAlbumName:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3ArtistName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadAlbumName:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGAlbumFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadYearName:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3AlbumName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadYearName:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGYearFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadComment:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3YearName(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadComment:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGCommentFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadGenre:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3Comment(1), HeaderByteString(FrameStartPos + TAGFramesStartPos + 4), MIN(FrameDataLength - 4, TAGSTRINGLENGTH)) 'add and subtract 4 (tested)
        End If
ReadGenre:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGGenreFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadComposer:
            End Select
            Tempstr$ = Mid$(HeaderString, FrameStartPos + TAGFramesStartPos, FrameDataLength)
            If HeaderByteString(FrameStartPos + TAGFramesStartPos) = 40 Then '(
                EndPos = InStr(1, Tempstr$, ")", vbBinaryCompare)
                If (EndPos) Then
                    '
                    'NOTE: there's an ID3v2 genre that has the following format:
                    '(0) Blues
                    '(12) Rock
                    '(125) Dance Hall.
                    '
                    MP3Genre = CByte(MIN(MAX(Val(Mid$(Tempstr$, 2, EndPos - 2 + 1)), 0), 255))
                    Tempstr$ = Trim$(Mid$(Tempstr$, EndPos + 1))
                    Call GETFIXEDBYTESTRINGFROMSTRING(TAGSTRINGLENGTH, MP3GenreName(), Tempstr$)
                Else
                    MP3Genre = 0 'reset (unknown)
                    Call GETFIXEDBYTESTRINGFROMSTRING(TAGSTRINGLENGTH, MP3GenreName(), Tempstr$)
                End If
            Else
                MP3Genre = 0 'reset (unknown)
                Call GETFIXEDBYTESTRINGFROMSTRING(TAGSTRINGLENGTH, MP3GenreName(), Tempstr$)
            End If
        End If
ReadComposer:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGComposerFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadOriginalArtist:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3Composer(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadOriginalArtist:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGOriginalArtistFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadCopyright:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3OriginalArtist(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadCopyright:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGCopyrightFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadURL:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3Copyright(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
ReadURL:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGURLFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo ReadEncodedBy:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3URL(1), HeaderByteString(FrameStartPos + TAGFramesStartPos + 1), MIN(FrameDataLength - 1, TAGSTRINGLENGTH)) 'add and subtract something (tested)
        End If
ReadEncodedBy:
        FrameStartPos = Frame_GetFramePos(HeaderString, TAGEncodedByFrameName)
        If (FrameStartPos) Then 'verify
            FrameDataLength = CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos)) - 1& 'no idea why we must subtract one
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 1&)) * (2& ^ 8&)
            FrameDataLength = FrameDataLength + CLng(HeaderByteString(FrameStartPos + FrameDataLengthEndPos - 2&)) * (2& ^ 16&)
            Select Case HeaderMajorVersion
            Case 3
                'skip frame if compressed or encrypted
                If Not ( _
                    ((HeaderByteString(FrameStartPos + 9) And 128) = 0) And _
                    ((HeaderByteString(FrameStartPos + 9) And 64) = 0)) Then GoTo Leave:
            End Select
            Call BYTESTRING_COPYMEMORY(MP3EncodedBy(1), HeaderByteString(FrameStartPos + TAGFramesStartPos), MIN(FrameDataLength, TAGSTRINGLENGTH))
        End If
Leave:
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3SongName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3ArtistName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3AlbumName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3YearName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3Comment(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3GenreName(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3Composer(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3OriginalArtist(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3Copyright(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3URL(), TAGSTRINGLENGTH)
        Call MP3TAG2_Read_VerifyMP3TAGItemByte(MP3EncodedBy(), TAGSTRINGLENGTH)
    Else
        MsgBox "internal error in MP3TAG2_ReadByteEx(): file " + Left$(MP3Name, 1024) + " not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    MP3TAG2_ReadByteEx = True 'ok
    Exit Function
Error: 'error or no TAG found
    Close #MP3NameFileNumber 'make sure file is closed
    MP3TAG2_ReadByteEx = False 'error
    Exit Function
End Function

'NOTE: the following code just didn't want to work right. I disabled all editing and
'just read & wrote the TAG. This worked. Then I enabled step-by-step writing each
'TAG item and were so able to corrected the tons of errors.
'NOTE: somehow this is one of the few situations where I don't know what I'm doing.
'But with a lot of hacking it will finally work!
'A quotation from the guy I sole the code from:
'"Skip over this frame...we do not know how to parse the frame data".
'NOTE: somehow people (including me) seem not to be sure how the frame data size
'is to be saved. In form of 8 bits per byte or only 7 bits.
'A sample I downloaded uses 7 bits (like in the TAG header), Winamp uses 8 bits
'(checked with hex editor). In the ID3v2.3 documentation there is no rule how to save
'the size (at least I didn't find one).

Public Function MP3TAG2_WriteByteEx(ByVal MP3Name As String, ByRef MP3SongName() As Byte, ByRef MP3ArtistName() As Byte, ByRef MP3AlbumName() As Byte, ByRef MP3YearName() As Byte, ByRef MP3Comment() As Byte, ByRef MP3GenreName() As Byte, ByRef MP3Genre As Byte, _
    ByRef MP3Composer() As Byte, ByRef MP3OriginalArtist() As Byte, ByRef MP3Copyright() As Byte, ByRef MP3URL() As Byte, ByRef MP3EncodedBy() As Byte) As Boolean
    On Error GoTo Error: 'important (if file locked, damaged or whatever)
    Dim MP3NameFileNumber As Integer
    Dim MP3LengthDelta As Long
    Dim HeaderLength As Double 'use double to avoid overflow error when calculating
    Dim HeaderLengthUnchanged As Long 'length of header in file before we manipulated something
    Dim HeaderMajorVersion As Byte
    Dim HeaderString As String
    Dim HeaderByteString() As Byte
    Dim DefaultHeaderUsedFlag As Boolean
    Dim FrameStartPos As Long
    Dim FrameLengthByte(0 To 3) As Byte
    Dim FrameDataLengthEndPos As Long 'points to last char of frame length (related to the complete frame but not to the complete TAG)
    Dim FrameDataLength As Long 'how many chars the data visible to the user has
    Dim FrameDataLengthOld As Long 'how many blah blah has had before overwriting
    Dim TAGFramesStartPos As Long
    Dim TAGHeaderLength As Long
    Dim TAGSongFrameName As String
    Dim TAGSongFrameDefault As String
    Dim TAGArtistFrameName As String
    Dim TAGArtistFrameDefault As String
    Dim TAGAlbumFrameName As String
    Dim TAGAlbumFrameDefault As String
    Dim TAGYearFrameName As String
    Dim TAGYearFrameDefault As String
    Dim TAGCommentFrameName As String
    Dim TAGCommentFrameDefault As String
    Dim TAGGenreFrameName As String
    Dim TAGGenreFrameDefault As String
    Dim TAGComposerFrameName As String
    Dim TAGComposerFrameDefault As String
    Dim TAGOriginalArtistFrameName As String
    Dim TAGOriginalArtistFrameDefault As String
    Dim TAGCopyrightFrameName As String
    Dim TAGCopyrightFrameDefault As String
    Dim TAGURLFrameName As String
    Dim TAGURLFrameDefault As String
    Dim TAGEncodedByFrameName As String
    Dim TAGEncodedByFrameDefault As String
    Dim EndPos As Long
    Dim Temp As Long
    Dim Tempstr$
    Dim TempByte As Byte
    Dim TempByteString() As Byte
    'Dim TempFile As String
    'Dim TempFileNumber As Integer
    '
    'NOTE: the code was mainly copied from MP3TAG2_WriteByteEx().
    'First we read the header string. When done, we take the string (not byte string)
    'and search for the frames. When the current frame to edit was found then we
    'cut the whole frame out of the header string and put in a new one.
    'If there isn't the current frame to edit then a default frame will be inserted and
    'then the searching is done again.
    'If the whole TAG header isn't existing then a default TAG header including all
    'frames that can be edited (got through the DefaultHeaderCreator) will be inserted.
    '
    'verify
    Call BYTESTRINGLEFT(MP3SongName(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3ArtistName(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3AlbumName(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3YearName(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3Comment(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3GenreName(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3Composer(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3OriginalArtist(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3URL(), 2& ^ 21&) 'code cannot write longer strings
    Call BYTESTRINGLEFT(MP3EncodedBy(), 2& ^ 21&) 'code cannot write longer strings
    'preset
    'TempFile = Stuff_GenerateTempFileName(STUFF_PROGRAMDIRECTORY)
    'If Len(TempFile) = 0 Then 'verify
    '    MsgBox "internal error in MP3TAG2_WriteByteEx(): temp file name could not be generated !", vbOKOnly + vbExclamation
    '    GoTo Error:
    'End If
    'TempFileNumber = FreeFile(0)
    'begin
    If (DirSave(MP3Name) = "") Or (Right$(MP3Name, 1) = "\") Or (MP3Name = "") Then 'verify
        MsgBox "internal error in MP3TAG2_WriteByteEx(): file '" + MP3Name + "' not found !", vbOKOnly + vbExclamation
        GoTo Error:
    End If
    MP3NameFileNumber = FreeFile(0)
    'Open TempFile For Output As #TempFileNumber
    'Close #TempFileNumber
    'read header string
    Open MP3Name For Binary As #MP3NameFileNumber
        'verify file length
        If LOF(MP3NameFileNumber) < 32 Then GoTo Error: 'there must be something in it to read first (11) header bytes
        'read header 'info block'
        ReDim TempByteString(1 To 10) As Byte
        Get #MP3NameFileNumber, 1, TempByteString()
        'check for string 'ID3'
        If Not ((TempByteString(1) = 73) Or (TempByteString(1) = 255)) Then
            GoSub CreateHeader: 'no 'I' or Chr$(255) (Chr$(255) really appeared in tests!)
        ElseIf Not (TempByteString(2) = 68) Then 'also check that, first char may be Chr$(255) by chance
            GoSub CreateHeader: 'no 'D'
        ElseIf Not (TempByteString(3) = 51) Then 'also check that, first char may be Chr$(255) by chance
            GoSub CreateHeader: 'no '3'
        End If
        'read header version
        HeaderMajorVersion = TempByteString(4)
        'byte 5: revision number
        'byte 6: flags
        'read header size
        HeaderLength = 0# 'reset
        HeaderLength = HeaderLength + (CDbl(TempByteString(7)) * 20917152#)
        HeaderLength = HeaderLength + (CDbl(TempByteString(8)) * 16384#)
        HeaderLength = HeaderLength + (CDbl(TempByteString(9)) * 128#)
        HeaderLength = HeaderLength + (CDbl(TempByteString(10)))
        If (HeaderLength < 1#) Or (HeaderLength > LOF(MP3NameFileNumber)) Or (HeaderLength > 2147483647#) Then 'verify
            GoTo Error:
        End If
        'save current original header length for later use
        If DefaultHeaderUsedFlag = False Then
            HeaderLengthUnchanged = CLng(HeaderLength)
        Else
            HeaderLengthUnchanged = 0
        End If
        'read header string (if not done yet)
        If DefaultHeaderUsedFlag = False Then
            HeaderString = String$(CLng(HeaderLength + 10), Chr$(0)) 'read file header, too
            Get #MP3NameFileNumber, 1, HeaderString 'read file header, too
            'NOTE: in HeaderString there's also the file header, but HeaderLength does not include the file header.
            Call GETBYTESTRINGFROMSTRING(Len(HeaderString), HeaderByteString(), HeaderString)
        Else
            'header string(s) already set
        End If
        '
    Close #MP3NameFileNumber
    'allocate header string
    '
    'NOTE: the blocks that store one tag item's data will be called the 'Frames'.
    'Every Frame is identified by its 'Frame name'.
    '
    Select Case HeaderMajorVersion 'two versions are supported 'no! (one version is supported)
'    Case 2 'ID3v2.2 'CAN'T BE WRITTEN (DOCUMENTATION MISSING)
'        TAGSongFrameName = "TT2"
'        TAGArtistFrameName = "TOA"
'        TAGAlbumFrameName = "TAL"
'        TAGYearFrameName = "TYE"
'        TAGCommentFrameName = "COM"
'        TAGGenreFrameName = "TCO"
'        TAGComposerFrameName = Chr$(0) 'not supported, Frame_GetFramePos() will notice the unsupported state and will not search for the frame
'        TAGOriginalArtistFrameName = Chr$(0) 'not supported, Frame_GetFramePos() will notice the unsupported state and will not search for the frame
'        TAGCopyrightFrameName = Chr$(0) 'not supported, Frame_GetFramePos() will notice the unsupported state and will not search for the frame
'        TAGURLFrameName = Chr$(0) 'not supported, Frame_GetFramePos() will notice the unsupported state and will not search for the frame
'        TAGEncodedByFrameName = Chr$(0) 'not supported, Frame_GetFramePos() will notice the unsupported state and will not search for the frame
'        TAGFramesStartPos = 7
'        TAGHeaderLength = TAGFramesStartPos - 1
'        FrameDataLengthEndPos = 5 'last char of the frame data length
    Case 3 'ID3v2.3
        TAGSongFrameName = "TIT2"
        TAGArtistFrameName = "TPE1"
        TAGAlbumFrameName = "TALB"
        TAGYearFrameName = "TYER"
        TAGCommentFrameName = "COMM"
        TAGGenreFrameName = "TCON"
        TAGComposerFrameName = "TCOM"
        TAGOriginalArtistFrameName = "TOPE"
        TAGCopyrightFrameName = "TCOP"
        TAGURLFrameName = "WXXX"
        TAGEncodedByFrameName = "TENC"
        TAGFramesStartPos = 11
        TAGHeaderLength = TAGFramesStartPos - 1
        FrameDataLengthEndPos = 7 'last char of the frame data length
    Case Else
        GoTo Error: 'header version is not supported
    End Select
    'set Artist-, artist-, album- etc. name
WriteSongName:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGSongFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteArtistName:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3SongName())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3SongName(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGSongFrameDefault = Chr$(84) + Chr$(73) + Chr$(84) + Chr$(50) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGSongFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteSongName:
        Else
            'song name can't be written :-(
        End If
    End If
WriteArtistName:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGArtistFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteAlbumName:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3ArtistName())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3ArtistName(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGArtistFrameDefault = Chr$(84) + Chr$(80) + Chr$(69) + Chr$(49) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGArtistFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteArtistName:
        Else
            'artist name can't be written :-(
        End If
    End If
WriteAlbumName:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGAlbumFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteYearName:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3AlbumName())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3AlbumName(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGAlbumFrameDefault = Chr$(84) + Chr$(65) + Chr$(76) + Chr$(66) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGAlbumFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteAlbumName:
        Else
            'album name can't be written :-(
        End If
    End If
WriteYearName:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGYearFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteComment:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3YearName())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3YearName(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGYearFrameDefault = Chr$(84) + Chr$(89) + Chr$(69) + Chr$(82) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGYearFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteYearName:
        Else
            'year name can't be written :-(
        End If
    End If
WriteComment:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGCommentFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteGenre:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3Comment())
        Call LongToByte(FrameDataLength + 1 + 4, FrameLengthByte()) 'add 1 plus 4 for any language-crap
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3Comment(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1 + 4) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld + 0) + 1) 'do NOT add four (where the + 0 is) for language ID or so-crap (tested)
    Else
        If HeaderMajorVersion = 3 Then
            TAGCommentFrameDefault = Chr$(67) + Chr$(79) + Chr$(77) + Chr$(77) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(5) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGCommentFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteComment:
        Else
            'comment can't be written :-(
        End If
    End If
WriteGenre:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGGenreFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteComposer:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        Call GETSTRINGFROMBYTESTRING(MP3GenreName(), Tempstr$)
        Tempstr$ = "(" + CStr(MP3Genre) + ")" + Tempstr$ '(255)UNKNOWN; don't use a space char between the bracket and the genre name, Winamp doesn't support this and also not recommended by ID3v2 documentation
        FrameDataLength = Len(Tempstr$)
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGGenreFrameDefault = Chr$(84) + Chr$(67) + Chr$(79) + Chr$(78) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGGenreFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteGenre:
        Else
            'genre name can't be written :-(
        End If
    End If
WriteComposer:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGComposerFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteOriginalArtist:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3Composer())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3Composer(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGComposerFrameDefault = Chr$(84) + Chr$(67) + Chr$(79) + Chr$(77) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGComposerFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteComposer:
        Else
            'composer can't be written :-(
        End If
    End If
WriteOriginalArtist:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGOriginalArtistFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteCopyright:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3OriginalArtist())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3OriginalArtist(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGOriginalArtistFrameDefault = Chr$(84) + Chr$(79) + Chr$(80) + Chr$(69) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGOriginalArtistFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteOriginalArtist:
        Else
            'original artist can't be written :-(
        End If
    End If
WriteCopyright:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGCopyrightFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteURL:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3Copyright())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3Copyright(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGCopyrightFrameDefault = Chr$(84) + Chr$(67) + Chr$(79) + Chr$(80) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGCopyrightFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteCopyright:
        Else
            'copyright can't be written :-(
        End If
    End If
WriteURL:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGURLFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo WriteEncodedBy:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3URL())
        Call LongToByte(FrameDataLength + 1 + 1, FrameLengthByte()) 'one extra for god-knows-what-it-is
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3URL(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1 + 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld + 0) + 1) 'do NOT add 1 (where the + 0 is) as the guys from ID3 stink
    Else
        If HeaderMajorVersion = 3 Then
            TAGURLFrameDefault = Chr$(87) + Chr$(88) + Chr$(88) + Chr$(88) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(2) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGURLFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteURL:
        Else
            'URL can't be written :-(
        End If
    End If
WriteEncodedBy:
    FrameStartPos = Frame_GetFramePos(HeaderString, TAGEncodedByFrameName)
    If (FrameStartPos) Then 'verify
        'skip writing if write-protected
        Select Case HeaderMajorVersion
        Case 3
            If (Asc(Mid$(HeaderString, FrameStartPos + 8, 1)) And 32) Then
                GoTo Leave:
            End If
        End Select
        'write data
        FrameDataLengthOld = _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2&, 1))) * 65536 + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1&, 1))) * 256& + _
            CLng(Asc(Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0&, 1))) - 1& 'not sure if we must use 7 or 8 bits per byte, we just do what Winamp 3 does
        FrameDataLength = GETBYTESTRINGLENGTH(MP3EncodedBy())
        Call LongToByte(FrameDataLength + 1, FrameLengthByte()) 'one must be added (tested) (for what reason ever, nothing found in documentation)
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 3, 1) = Chr$(0) 'longer length not supported
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 2, 1) = Chr$(FrameLengthByte(1))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 1, 1) = Chr$(FrameLengthByte(2))
        Mid$(HeaderString, FrameStartPos + FrameDataLengthEndPos - 0, 1) = Chr$(FrameLengthByte(3))
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 64) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 64) 'clear encryption flag
        If (Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) And 128) Then Mid$(HeaderString, FrameStartPos + 9, 1) = Chr$(Asc(Mid$(HeaderString, FrameStartPos + 9, 1)) Xor 128) 'clear compression flag
        Call GETSTRINGFROMBYTESTRING(MP3EncodedBy(), Tempstr$)
        HeaderString = _
            Left$(HeaderString, FrameStartPos + TAGFramesStartPos - 1) + _
            Tempstr$ + _
            Right$(HeaderString, Len(HeaderString) - (FrameStartPos + TAGFramesStartPos + FrameDataLengthOld) + 1)
    Else
        If HeaderMajorVersion = 3 Then
            TAGEncodedByFrameDefault = Chr$(84) + Chr$(69) + Chr$(78) + Chr$(67) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(64) + Chr$(0) + Chr$(0)
            HeaderString = _
                Left$(HeaderString, TAGHeaderLength) + _
                TAGEncodedByFrameDefault + _
                Mid$(HeaderString, TAGFramesStartPos)
            GoTo WriteEncodedBy:
        Else
            'encoded by can't be written :-(
        End If
    End If
Leave:
    'manipulate header
    '
    'NOTE: make sure header's size must be divisable through 257. This is for some reason
    'necessary, we were told so in a downloaded sample and expired it through tests.
    '
    Call GETBYTESTRINGFROMSTRING(Len(HeaderString), TempByteString(), HeaderString)
    For Temp = UBound(TempByteString()) To 16 Step (-1) 'does have a minimal length, verified when reading original header
        '
        'NOTE: do not just add null chars, try to use existing ones as far as possible
        'to avoid that the mp3 file increases its size with every ID3v2 TAG writing.
        'NOTE: a header string could look the following:
        'TCOPxxxxff00000000000000000000000000000000000...
        'The flag bytes or frame data length bytes could also be 0, don't cut them,
        'leave for safty reasons 16 bytes of null chars standing behind the last frame
        'identifier or frame flags or frame data (just behind the last non-null char).
        '
        If Not (TempByteString(Temp - 16) = 0) Then
            HeaderString = Left$(HeaderString, Temp)
            Exit For
        End If
    Next Temp
    '
    HeaderLength = Len(HeaderString) - TAGHeaderLength 'exclude file header and 'fill null chars' from header length
    HeaderString = HeaderString + String$( _
        257 - (HeaderLength Mod 257), Chr$(0))
    HeaderLength = Len(HeaderString) - TAGHeaderLength 'exclude file header and 'fill null chars' from header length
    'write total header size
    '
    'NOTE: the first bit of every written byte is cleared,
    'so we write 7 bits per byte only.
    '
    Mid$(HeaderString, 7, 1) = Chr$(Int(HeaderLength / 20917152#))
    HeaderLength = HeaderLength - (Int(HeaderLength / 20917152#) * 20917152#)
    Mid$(HeaderString, 8, 1) = Chr$(Int(HeaderLength / 16384#))
    HeaderLength = HeaderLength - (Int(HeaderLength / 16384#) * 16384#)
    Mid$(HeaderString, 9, 1) = Chr$(Int(HeaderLength / 128#))
    HeaderLength = HeaderLength - (Int(HeaderLength / 128#) * 128#)
    Mid$(HeaderString, 10, 1) = Chr$(HeaderLength)
    HeaderLength = Len(HeaderString) - TAGHeaderLength 'reset
    'now open the original mp3 file, shrink or enlarge it and move audio data
    MP3LengthDelta = HeaderLengthUnchanged - HeaderLength 'the 10 chars of the TAG header were subtracted from both values
    Select Case MP3LengthDelta
    Case Is < 0
        Call GFEnlargeFile(MP3Name, FileLen(MP3Name) + MP3LengthDelta)
        GoSub MoveAudioData:
    Case Is > 0
        GoSub MoveAudioData:
        Call GFShrinkFile(MP3Name, FileLen(MP3Name) + MP3LengthDelta)
    Case Is = 0
        'nothing to move/enlarge/shrink
    End Select
    Open MP3Name For Binary As #MP3NameFileNumber
        Put #MP3NameFileNumber, 1, HeaderString
    Close #MP3NameFileNumber
    MP3TAG2_WriteByteEx = True 'ok
    Exit Function
Error:
    'Close #TempFileNumber
    Close #MP3NameFileNumber
    MP3TAG2_WriteByteEx = False 'error
    Exit Function
MoveAudioData:
    Dim BlockStartPos As Long 'start pos of block to move
    Dim BlockLength As Long 'length of block to move
    Dim BlockString As String
    'preset
    Select Case MP3LengthDelta
    Case Is < 0
        BlockStartPos = FileLen(MP3Name) - 8192000 'first copy 1 byte, then the rest
    Case Is > 0
        BlockStartPos = 1 + HeaderLengthUnchanged
    End Select
    'begin
    Open MP3Name For Binary As #MP3NameFileNumber
        Do
            BlockLength = 8192000 'read approx. 8 MB at once
            Select Case MP3LengthDelta
            Case Is < 0
                If (BlockStartPos + BlockLength - 1) > LOF(MP3NameFileNumber) Then
                    BlockLength = LOF(MP3NameFileNumber) - BlockStartPos + 1
                End If
                If BlockLength = 0 Then Exit Do 'should not happen (but avoid endless loop)
                If BlockStartPos < (1 + TAGHeaderLength) Then Exit Do 'verify
            Case Is > 0
                If (BlockStartPos + BlockLength - 1) > LOF(MP3NameFileNumber) Then
                    BlockLength = LOF(MP3NameFileNumber) - BlockStartPos + 1
                End If
                If BlockLength = 0 Then Exit Do 'verify
            End Select
            BlockString = String$(BlockLength, Chr$(0))
            Get #MP3NameFileNumber, BlockStartPos, BlockString
            Put #MP3NameFileNumber, BlockStartPos + (-MP3LengthDelta), BlockString
            Select Case MP3LengthDelta
            Case Is < 0
                BlockStartPos = BlockStartPos - BlockLength
            Case Is > 0
                BlockStartPos = BlockStartPos + BlockLength
            End Select
        Loop
    Close #MP3NameFileNumber
    Return
CreateHeader:
    DefaultHeaderUsedFlag = True 'do not read the original header string (as there isn't one)
    HeaderString = _
        Chr$(73) + Chr$(68) + Chr$(51) + Chr$(3) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(5) + Chr$(45) + Chr$(84) + Chr$(82) + Chr$(67) + Chr$(75) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(69) + Chr$(78) + Chr$(67) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(64) + Chr$(0) + Chr$(0) + _
        Chr$(87) + Chr$(88) + Chr$(88) + Chr$(88) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(2) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(67) + Chr$(79) + Chr$(80) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(79) + Chr$(80) + Chr$(69) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(84) + Chr$(67) + Chr$(79) + Chr$(77) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(67) + Chr$(79) + _
        Chr$(78) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(67) + Chr$(79) + Chr$(77) + Chr$(77) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(5) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(89) + Chr$(69) + Chr$(82) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(84) + Chr$(65) + Chr$(76) + Chr$(66) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(80) + Chr$(69) + _
        Chr$(49) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(84) + Chr$(73) + Chr$(84) + Chr$(50) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(1) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + _
        Chr$(0) 'got through using DefaultHeaderCreator (in current directory)
    Call GETBYTESTRINGFROMSTRING(10, TempByteString(), HeaderString)
    Call GETBYTESTRINGFROMSTRING(Len(HeaderString), HeaderByteString(), HeaderString)
    Return
End Function

'***FRAME FUNCTIONS***
'NOTE: the following function are used to handle ID3v2 TAG frames.

Private Function Frame_GetFramePos(ByRef HeaderString As String, ByVal FrameIdentifier As String) As Long
    'on error resume next 'returns frame start pos or 0 for frame not found
    Dim SearchStartPos As Long
    '
    'NOTE: if we manage to determinate all frame positions through 'walking'
    'through the frames then store every frame's pos and identifier in a table (array).
    'This function then must search for the identifier in the table and return the
    'related start pos. But until now we couldn't manage to deteminate the frame
    'positions in an other way than using InStr().
    '
    'verify
    If Asc(FrameIdentifier) = 0 Then
        Frame_GetFramePos = 0 'frame not supported
        Exit Function
    End If
    'preset
    SearchStartPos = 1
    'begin
ReDo:
    Frame_GetFramePos = InStr(SearchStartPos, HeaderString, FrameIdentifier, vbBinaryCompare)
    If (Frame_GetFramePos) Then
        '
        'NOTE: if a position was found, then check if the char behind the frame header
        'is a Chr$(0), as there is the fourth byte of the frame data size, which
        'should be 0, if a stupid user entered frame data longer than approx.
        '256 ^ 3 chars then the frame will not be found (blame the guys from ID3!).
        '
        If Not ((Frame_GetFramePos + 4) > Len(HeaderString)) Then 'verify
            If Asc(Mid$(HeaderString, (Frame_GetFramePos + 4), 1)) = 0 Then
                'ok
            Else
                SearchStartPos = (Frame_GetFramePos + 4) 'search for next appearence of frame identifier (if any)
                Frame_GetFramePos = 0 'probably not a frame header but e.g. 'TV COMMERCIAL (The Album)' (COMM = comments frame identifier), located before comments string
                GoTo ReDo:
            End If
        Else
            Frame_GetFramePos = 0 'there cannot be a valid frame anyway
        End If
    End If
    Exit Function
End Function

Private Function Frame_GetNextFrameString(ByRef HeaderString As String, ByVal HeaderLengthTotal As Long, ByRef FrameStartPos As Long, ByRef FrameString As String, ByRef FrameDescription As String) As Boolean 'HeaderString passed ByRef to increase speed
    'on error resume next 'returns True if frame string found, False if not
    '
    'NOTE: FrameStartPos must point to the start pos of the first frame (11) when calling
    'this function the first time. The whole string of the next frame will be copied to
    'FrameString. FrameStartPos will be increased, pass this value when next time
    'calling this function. FrameDescription is set to the first 4 chars of the FrameString.
    '
    'begin
    FrameDescription = Left$(HeaderString, 4)
    Select Case FrameDescription
    Case ""
    '
    '**********************************************************************
    'STOP!!! DAMN!! These ID3 guys!!!
    'Were they drunken when designing the ID3v2 TAG format?
    'I don't know the frame header size for every frame (varies),
    'so where does the next frame start???
    'These ID3 guys!!
    'We must use InStr() to find the frames :-(
    'This function cannot be used.
    '**********************************************************************
    '
    End Select
End Function

'***END OF FRAME FUNCTIONS***
'***********************************END OF MP3 TAG 2***********************************
'***********************************GENERAL FUNCTIONS**********************************

Private Function GFShrinkFile(ByVal ShrinkName As String, ByVal ShrinkNameSizeNew As Long) As Boolean
    'on error Resume Next 'shrinks a file; function returns True if file was shrinked, False if not
    Dim ShrinkNameHandle As Long
    Dim OFSTRUCTVar As OFSTRUCT
    Dim ShrinkFileTemp As Long
    'verify
    If ((Dir(ShrinkName) = "") Or (Right$(ShrinkName, 1) = "\") Or (ShrinkName = "")) Then 'verify
        GFShrinkFile = False 'error
        Exit Function
    End If
    Select Case ShrinkNameSizeNew
    Case Is < 0
        GoTo Error:
    Case Is > FileLen(ShrinkName)
        ShrinkNameSizeNew = FileLen(ShrinkName)
    End Select
    'begin
    ShrinkNameHandle = OpenFile(ShrinkName, OFSTRUCTVar, OF_READWRITE)
    If ShrinkNameHandle = 0 Then GoTo Error: 'verify
    ShrinkFileTemp = SetFilePointer(ShrinkNameHandle, ShrinkNameSizeNew, 0, FILE_BEGIN)
    'If ShrinkFileTemp = 0 Then GoTo Error: 'functions returns something nobody understands
    ShrinkFileTemp = SetEndOfFile(ShrinkNameHandle)
    If ShrinkFileTemp = 0 Then GoTo Error: 'verify
    ShrinkFileTemp = CloseHandle(ShrinkNameHandle)
    If ShrinkFileTemp = 0 Then GoTo Error: 'verify
    GFShrinkFile = True 'ok
    Exit Function
Error:
    Call CloseHandle(ShrinkNameHandle) 'make sure file is closed
    GFShrinkFile = False 'error
    Exit Function
End Function

Private Function GFEnlargeFile(ByVal EnlargeName As String, ByVal EnlargeNameSizeNew As Long) As Boolean
    'on error resume next 'enlarges a file; function returns True if file was enlarged, False if not
    Dim EnlargeNameHandle As Long
    Dim OFSTRUCTVar As OFSTRUCT
    Dim EnlargeFileTemp As Long
    'verify
    If ((Dir(EnlargeName) = "") Or (Right$(EnlargeName, 1) = "\") Or (EnlargeName = "")) Then 'verify
        GFEnlargeFile = False 'error
        Exit Function
    End If
    Select Case EnlargeNameSizeNew
    Case Is < FileLen(EnlargeName)
        GoTo Error:
    'Case Is > (2# ^ 32# - 2#) 'VC++ help 'EnlargeNameSizeNew has the type Long anyway
    '    EnlargeNameSizeNew = (2# ^ 32# - 2#)
    End Select
    'begin
    EnlargeNameHandle = OpenFile(EnlargeName, OFSTRUCTVar, OF_READWRITE)
    If EnlargeNameHandle = 0 Then GoTo Error: 'verify
    EnlargeFileTemp = SetFilePointer(EnlargeNameHandle, EnlargeNameSizeNew, 0, FILE_BEGIN)
    'If EnlargeFileTemp = 0 Then GoTo Error: 'functions returns something nobody understands
    EnlargeFileTemp = SetEndOfFile(EnlargeNameHandle)
    If EnlargeFileTemp = 0 Then GoTo Error: 'verify
    EnlargeFileTemp = CloseHandle(EnlargeNameHandle)
    If EnlargeFileTemp = 0 Then GoTo Error: 'verify
    GFEnlargeFile = True 'ok
    Exit Function
Error:
    Call CloseHandle(EnlargeNameHandle) 'make sure file is closed
    GFEnlargeFile = False 'error
    Exit Function
End Function

'*******************************END OF GENERAL FUNCTIONS*******************************
'*************************************COPIED CODE**************************************
'NOTE: the following code was not written by me, but just manipulated so
'that it looks a bit like my style (hehehe).

Private Sub LongToByte(ByVal Val As Long, ByRef ByteArray() As Byte)
    On Error GoTo ErrHandler:
    
'    GS07312001 -   Replaced with the Correct Implementation
'    Dim byte1 As Byte
'    Dim byte2 As Byte
'    Dim byte3 As Byte
'    Dim byte4 As Byte
'
'    byte1 = val And 127
'    val = val / 128
'
'    byte2 = val And 127
'    val = val / 128
'
'    byte3 = val And 127
'    val = val / 128
'
'    byte4 = val And 127
'
'    ReDim byteArray(3)
'    byteArray(0) = byte4
'    byteArray(1) = byte3
'    byteArray(2) = byte2
'    byteArray(3) = byte1

    Dim idx As Long

    'NOTE: manipulated by Louis, orignal line was:
    'ByteArray(idx) = ShiftBits(Val, (3& - idx) * 7&, eShiftRight) And 127&
    
    For idx = 0 To 3
        ByteArray(idx) = ShiftBits(Val, (3& - idx) * 8&, eShiftRight) And 255&
    Next idx

    On Error GoTo 0
    Exit Sub

ErrHandler:
    'Raise the error back to the caller
    Err.Raise Err.Number, "ID3v2Enums::LongToByte", Err.Description
    Exit Sub
End Sub

Private Function ShiftBits(ByVal lValue As Long, ByVal lNumBitsToShift As Long, ByVal eDir As eBitShiftDir) As Long
    On Error GoTo ErrHandler:
   
    Select Case eDir
    Case eShiftLeft
        ShiftBits = lValue * (2& ^ lNumBitsToShift)
    Case eShiftRight
        ShiftBits = lValue \ (2& ^ lNumBitsToShift)
    End Select
    
    On Error GoTo 0
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "ShiftBits", Err.Description
    Exit Function
End Function

'*********************************END OF COPIED CODE***********************************
'****************************************OTHER*****************************************

Private Function DirSave(ByRef PathName As String, Optional ByVal Attributes As Integer = vbNormal) As String
    'on error resume next
    '
    'NOTE: copied as the copied MP3TAG1 code uses DirSave().
    '
    'begin
    DirSave = GFFileAccess_DirSave(PathName, Attributes)
End Function

Private Sub RemoveChr0(ByRef RemoveString As String)
    'on error Resume Next 'replaces Chr$(0) through space
    Dim Temp As Long
    If InStr(1, RemoveString, Chr$(0), vbBinaryCompare) = 0 Then Exit Sub 'nothing to remove
    For Temp = 1 To Len(RemoveString)
        If Asc(Mid$(RemoveString, Temp, 1)) = 0 Then
            Mid$(RemoveString, Temp, 1) = " "
        End If
    Next Temp
End Sub

Private Sub RemoveChr0Byte(ByRef RemoveByteString() As Byte, ByRef RemoveByteStringLength As Long)
    'on error Resume Next 'replaces Chr$(0) through space
    Dim Temp As Long
    For Temp = 1 To RemoveByteStringLength
        If RemoveByteString(Temp) = 0 Then RemoveByteString(Temp) = 32
    Next Temp
End Sub

Private Function MIN(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use for i.e. CopyMemory(a(1), ByVal b, MIN(UBound(a()), Len(b))
    If Value1 < Value2 Then
        MIN = Value1
    Else
        MIN = Value2
    End If
End Function

Private Function MAX(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    'on error Resume Next 'use in combination with ReDim()
    If Value1 > Value2 Then
        MAX = Value1
    Else
        MAX = Value2
    End If
End Function

