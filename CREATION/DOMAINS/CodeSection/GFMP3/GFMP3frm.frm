VERSION 5.00
Begin VB.Form GFMP3frm 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command6 
      Caption         =   "MP3TAG1_Write()"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MP3TAG1_Read()"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   900
      Width           =   2415
   End
   Begin VB.CommandButton MP3NameBrowseCommand 
      Caption         =   "..."
      Height          =   315
      Left            =   4020
      TabIndex        =   1
      Top             =   60
      Width           =   555
   End
   Begin VB.TextBox MP3NameText 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "[select mp3 file]"
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy"
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox WriteTAGDataText 
      Height          =   266
      Index           =   0
      Left            =   56
      TabIndex        =   2
      Top             =   56
      Width           =   1918
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MP3TAG2_WriteByteEx()"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MP3TAG2_ReadByteEx()"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2340
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MP3TAG2_ReadByte()"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
End
Attribute VB_Name = "GFMP3frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001-2003, 2004,2010 by Louis.
'
'Downloaded from www.louis-coder.com.
'Add all modules of this project to your target project.
'With the functions in GFMP3mod you can read and write ID3v1 and v2 tags.
'Please note that I cannot guarantee that the ID3v2 tag reading and writing is
'100% correct, because the standard is such a mess. For instance, nobody
'(I asked a LOT of people) really knows how many bits are used to describe
'the frame sizes (containing the string information), the original designer of
'the ID3v2 standard didn't reply to my mails.
'I downloaded several samples, and found out that Winamp 3 stores the tag
'data differently than other editors, even ignoring standard rules.
'But the code should work at least (!) as well as other samples out there.
'
'To speed up the reading and writing of many tags, 'byte strings' are used,
'that means strings aren't saved in slow VB strings but in byte arrays
'(like in C). To convert between byte strings and VB strings you should use
'the BSmod functions (see sample code below).
'A special version of the GFMP3 code was used in Toricxs (www.toricxs.com).
'For questions or comments you can mail louis@louis-coder.com.
'
Dim GenreNameArray As Variant

Private Sub Form_Load()
    'on error resume next
    'Debug.Print STUFF_GETBYTEBITWINDOW(RGB(12, 123, 234), 1, 8) 'first implemented for this project (DEBUG)
    'Debug.Print STUFF_GETBYTEBITWINDOW(RGB(12, 123, 234), 9, 8) 'DEBUG
    'Debug.Print STUFF_GETBYTEBITWINDOW(RGB(12, 123, 234), 17, 8) 'DEBUG
    Dim Temp As Long
    For Temp = 0 To 10
        If Temp > 0 Then
            Load WriteTAGDataText(Temp)
            WriteTAGDataText(Temp).Top = WriteTAGDataText(0).Top + 20 * Screen.TwipsPerPixelY * Temp
            WriteTAGDataText(Temp).Visible = True
        End If
        Select Case Temp
        Case 0
            WriteTAGDataText(Temp).ToolTipText = "song name"
        Case 1
            WriteTAGDataText(Temp).ToolTipText = "artist name"
        Case 2
            WriteTAGDataText(Temp).ToolTipText = "album name"
        Case 3
            WriteTAGDataText(Temp).ToolTipText = "year"
        Case 4
            WriteTAGDataText(Temp).ToolTipText = "comment"
        Case 5
            WriteTAGDataText(Temp).ToolTipText = "genre name"
        Case 6
            WriteTAGDataText(Temp).ToolTipText = "composer"
        Case 7
            WriteTAGDataText(Temp).ToolTipText = "original artist"
        Case 8
            WriteTAGDataText(Temp).ToolTipText = "copyright"
        Case 9
            WriteTAGDataText(Temp).ToolTipText = "url"
        Case 10
            WriteTAGDataText(Temp).ToolTipText = "encoded by"
        End Select
    Next Temp
    GenreNameArray = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", _
        "Hip-Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&b", "Rap", "Reggae", _
        "Rock", "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", _
        "Soundtrack", "Euro - Techno", "Ambient", "Trip-Hop", "Vocal", "JazzFunk", "Fusion", _
        "Trance", "Classical", "Instrumental", "Acid", "House", "Game", "SoundClip", "Gospel", _
        "Noise", "AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
        "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial", "Electronic", _
        "Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", "Top 40", _
        "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", "New Wave", _
        "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", "Acid Punk", "Acid Jazz", _
        "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", "Folk", "Folk/Rock", "National Folk", _
        "Swing", "Bebob", "Latin", "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", _
        "Progressive Rock", "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", _
        "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty Bass", _
        "Primus", "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", "Power Ballad", _
        "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Cappella", _
        "Euro - House", "Dance Hall", "Goa", "Drum & Bass", "Club - House", "Hardcore", "Terror", "Indie", "BritPop", _
        "Negerpunk", "Polsk Punk", "Beat", "Christian Gangsta Rap", "Heavy Metal", "Black Metal", "Crossover", _
        "Contemporary Christian", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", "Anime", "JPop", "Synthpop")
End Sub

Private Sub MP3NameBrowseCommand_Click()
    'on error resume next
    Dim Tempstr$
    Tempstr$ = Stuff_GFCDGetFileNameFast("Select mp3 file for TAG edit", App.Path)
    If (Len(Tempstr$)) Then MP3NameText.Text = Tempstr$
End Sub

Private Sub Command1_Click()
    'on error resume next
    Dim a(1 To 1024) As Byte
    Dim b(1 To 1024) As Byte
    Dim c(1 To 1024) As Byte
    Dim d(1 To 1024) As Byte
    Dim e(1 To 1024) As Byte
    Dim f(1 To 1024) As Byte
    Dim g As Byte
    Dim s As Single
    Dim Temp As Long
    'begin
    's = Timer
    'For Temp = 1 To 1000
        Debug.Print MP3TAG2_ReadByte(MP3NameText.Text, _
            a(), b(), c(), d(), e(), f(), g, 1024)
    'Next Temp
    'Debug.Print Timer - s
    Debug.Print "ID3V2 TAG: (at " + Time$ + ")"
    Call DISPLAYBYTESTRING(a())
    Call DISPLAYBYTESTRING(b())
    Call DISPLAYBYTESTRING(c())
    Call DISPLAYBYTESTRING(d())
    Call DISPLAYBYTESTRING(e())
    Call DISPLAYBYTESTRING(f())
    Debug.Print g
End Sub

Private Sub Command2_Click()
    'on error resume next
    Dim a(1 To 1024) As Byte
    Dim b(1 To 1024) As Byte
    Dim c(1 To 1024) As Byte
    Dim d(1 To 1024) As Byte
    Dim e(1 To 1024) As Byte
    Dim f(1 To 1024) As Byte
    Dim g As Byte
    Dim h(1 To 1024) As Byte
    Dim i(1 To 1024) As Byte
    Dim j(1 To 1024) As Byte
    Dim k(1 To 1024) As Byte
    Dim l(1 To 1024) As Byte
    Dim s As Single
    Dim Temp As Long
    'begin
    's = Timer
    'For Temp = 1 To 1000
        Debug.Print MP3TAG2_ReadByteEx(MP3NameText.Text, _
            a(), b(), c(), d(), e(), f(), g, h(), i(), j(), k(), l(), 1024)
    'Next Temp
    'Debug.Print Timer - s
    Debug.Print "ID3V2 TAG: (at " + Time$ + ")"
    WriteTAGDataText(0).Text = GETRETURNSTRINGFROMBYTESTRING(a())
    WriteTAGDataText(1).Text = GETRETURNSTRINGFROMBYTESTRING(b())
    WriteTAGDataText(2).Text = GETRETURNSTRINGFROMBYTESTRING(c())
    WriteTAGDataText(3).Text = GETRETURNSTRINGFROMBYTESTRING(d())
    WriteTAGDataText(4).Text = GETRETURNSTRINGFROMBYTESTRING(e())
    WriteTAGDataText(5).Text = GETRETURNSTRINGFROMBYTESTRING(f())
    WriteTAGDataText(6).Text = GETRETURNSTRINGFROMBYTESTRING(h())
    WriteTAGDataText(7).Text = GETRETURNSTRINGFROMBYTESTRING(i())
    WriteTAGDataText(8).Text = GETRETURNSTRINGFROMBYTESTRING(j())
    WriteTAGDataText(9).Text = GETRETURNSTRINGFROMBYTESTRING(k())
    WriteTAGDataText(10).Text = GETRETURNSTRINGFROMBYTESTRING(l())
    Debug.Print g 'genre (number)
End Sub

Private Sub Command3_Click()
    'on error resume next
    Dim a(1 To 1024) As Byte
    Dim b(1 To 1024) As Byte
    Dim c(1 To 1024) As Byte
    Dim d(1 To 1024) As Byte
    Dim e(1 To 1024) As Byte
    Dim f(1 To 1024) As Byte
    Dim g As Byte
    Dim h(1 To 1024) As Byte
    Dim i(1 To 1024) As Byte
    Dim j(1 To 1024) As Byte
    Dim k(1 To 1024) As Byte
    Dim l(1 To 1024) As Byte
    'preset
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, a(), WriteTAGDataText(0).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, b(), WriteTAGDataText(1).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, c(), WriteTAGDataText(2).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, d(), WriteTAGDataText(3).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, e(), WriteTAGDataText(4).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, f(), WriteTAGDataText(5).Text)
    g = 255 'genre
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, h(), WriteTAGDataText(6).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, i(), WriteTAGDataText(7).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, j(), WriteTAGDataText(8).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, k(), WriteTAGDataText(9).Text)
    Call GETFIXEDBYTESTRINGFROMSTRING(1024, l(), WriteTAGDataText(10).Text)
    'begin
    Debug.Print MP3TAG2_WriteByteEx(MP3NameText.Text, _
        a(), b(), c(), d(), e(), f(), g, h(), i(), j(), k(), l())
End Sub

Private Sub Command4_Click()
    'on error resume next
    'Call FileCopy("C:\Music Instructor - Get Freaky.mp3", "C:\Music Instructor - Get Freaky (Test).mp3") 'DEBUG
End Sub

Private Sub Command5_Click()
    'on error resume next
    Dim s1 As String, s2 As String, s3 As String, s4 As String, s5 As String
    Dim Genre As Byte
    Call MP3TAG1_Read(MP3NameText.Text, s1, s2, s3, s4, s5, Genre)
    WriteTAGDataText(0).Text = s1
    WriteTAGDataText(1).Text = s2
    WriteTAGDataText(2).Text = s3
    WriteTAGDataText(3).Text = s4
    WriteTAGDataText(4).Text = s5
    If (Genre < 146) Then
        WriteTAGDataText(5).Text = GenreNameArray(Genre)
    Else
        WriteTAGDataText(5).Text = "[unknown genre]"
    End If
End Sub

Private Sub Command6_Click()
    'on error resume next
    Dim s1 As String, s2 As String, s3 As String, s4 As String, s5 As String
    Dim Genre As Byte
    'Genre = GetGenreNameFromList(GenreNameArray) 'make the user select the genre from a list (not implemented here)
    Call MP3TAG1_Write(MP3NameText.Text, WriteTAGDataText(0).Text, WriteTAGDataText(1).Text, WriteTAGDataText(2).Text, WriteTAGDataText(3).Text, WriteTAGDataText(4).Text, Genre)
End Sub

