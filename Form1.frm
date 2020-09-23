VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirBox Sample"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000003&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000003&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "©SoulSeeker"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************'
'  Coded by Ivan Zlatev (programs@mail.bg)  '
'  You can use this code whenever you want !'
'                                 cya!      '
'*******************************************'

'Some Declarations...
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
hWndOwner      As Long
pIDLRoot       As Long
pszDisplayName As Long
lpszTitle      As Long
ulFlags        As Long
lpfnCallback   As Long
lParam         As Long
iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Dim kbits As Variant



Private Sub Command1_Click()
' Here is the title of the DirBox

DirBox "Here is the DirBox...", id$

' And here is where will be the browser destination displayed

Text1.Text = id$




End Sub
Private Sub DirBox(Msg As String, Directory As String)
On Error Resume Next
    
'Well, i will say change this when you know what do you do :).
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    'Change this to set what info is displayed.
    szTitle = Msg
    With tBrowseInfo
       .hWndOwner = Me.hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       Directory = sBuffer
    End If
End Sub


Private Sub Command2_Click()
On Error Resume Next
' Here is the action of the 'Exit' button.This exits the program.

Unload Me    'You can use 'End',too.

End Sub

Private Sub Command3_Click()
On Error Resume Next
' Here is my info.
MsgBox "By SoulSeeker" & vbNewLine & "programs@mail.bg", vbInformation
End Sub
