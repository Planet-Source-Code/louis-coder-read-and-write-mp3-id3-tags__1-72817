VERSION 5.00
Begin VB.Form Testfrm 
   Caption         =   "GGFFileAccess Test"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "GetFreeDiskSpace"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GetTotalDiskSpace"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Text            =   "C:\"
      Top             =   840
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetDirFileSizeTotal"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "Testfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2001, 2004 by Louis. Test form for GFFileAccess.
'
'Downloaded from www.louis-coder.com.
'GFFileAccessmod can be used for:
'-determinating the size of all files in a dir EXTREMELY FAST;
'-determinating the free and total disk space of a drive
' (also works on Win95 OSR 1).

Private Sub Command1_Click()
    'on error resume next
    Debug.Print GFFileAccess_GetDirFileSizeTotal(Text1.Text, Text2.Text)
End Sub

Private Sub Command2_Click()
    'on error resume next
    Debug.Print GFFileAccess_GetTotalDiskSpace(Text3.Text)
End Sub

Private Sub Command3_Click()
    'on error resume next
    Debug.Print GFFileAccess_GetFreeDiskSpace(Text3.Text)
End Sub

