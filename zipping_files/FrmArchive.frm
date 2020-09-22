VERSION 5.00
Begin VB.Form FrmArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zipping Files"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   Icon            =   "FrmArchive.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox List1 
      Height          =   3795
      Left            =   1800
      TabIndex        =   9
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create Zip"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zip Saving Path"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Folder"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Files to be zip:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Output Zip File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "FrmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MakeZip()
'On Error Resume Next
sZipFile = ""
Dim xDate As Data
If List1.ListCount > 0 Then
Dim i As Integer
Dim SSSTATUS As Integer
Dim ssczip As ZipClass
   Set ssczip = New ZipClass
    For i = 0 To List1.ListCount - 1
         List1.ListIndex = i
         On Error Resume Next
         ssczip.AddFile Text1.Text & "\" & List1.FileName
    Next
   sZipFile = Text2.Text '& ".zip"
   ssczip.WriteZip sZipFile, True

   Set ssczip = Nothing
   MsgBox "Folder contants has been compressed in " & Text2.Text
   Else
   MsgBox "Nothing for making zip archive!!!"
End If

End Sub


Private Sub Command2_Click()
   Dim BrFolder As String
  BrFolder = BrowseForFolder("Select for Folder:")
  If BrFolder <> "" Then
    Text1.Text = BrFolder
    List1.Path = Text1.Text
  End If
End Sub

Private Sub Command3_Click()
   Dim BrFolder As String
  BrFolder = BrowseForFolder("Select for Folder:")
  If BrFolder <> "" Then
    Text2.Text = BrFolder & "\" & "New File.zip"
  End If
End Sub

Private Sub Command4_Click()
If Text2.Text = "" Then
MsgBox "Please enter the zip file path"
Exit Sub
End If
If Mid(Text2.Text, Len(Text2.Text) - 2) = "zip" Then
Call MakeZip
Else
Text2.Text = Text2.Text & ".zip"
Call MakeZip
End If
End Sub

Private Sub Form_Load()
List1.Refresh
End Sub
