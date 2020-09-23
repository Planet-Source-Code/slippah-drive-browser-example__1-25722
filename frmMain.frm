VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Browser"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox DriveBox 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
   Begin VB.FileListBox FileBox 
      Height          =   4185
      Left            =   2640
      TabIndex        =   1
      Top             =   350
      Width           =   2655
   End
   Begin VB.DirListBox DirBox 
      Height          =   4140
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblFileName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   5260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was written by Slippah, that is, all except the "Error" section
'of Drivebox_Change. This program acts as an alternative to Windows
'Explorer, or My Computer. Unlike the recent upload to PSC, File Browser,
'this program allows you to open files from within the application! This
'can be accomplished by double-clicking on the file that you wish to open
'in the FileBox. For some reason, this does not ALWAYS work, but it works
'enough, I guess, and it adds a dimension to the program not found in many
'other applications on PSC.
Private Sub Dirbox_Change()
FileBox.Path = DirBox.Path
End Sub

Private Sub Drivebox_Change()
Dim msg As String
Dim answer As String
On Error GoTo Error
DirBox.Path = DriveBox.Drive
Exit Sub
'The section below this text, the Error section I didn't write
Error:
msg = "Error: " & Err.Number & ": " & Err.Description
answer = MsgBox(msg, vbOKCancel, "Error in setting path")
If answer = vbOK Then
Resume
Else
DriveBox.Drive = DirBox.Path
Err.Clear
Exit Sub
End If
End Sub

Private Sub Filebox_Click()
lblFileName.Caption = DirBox.Path + "\" + FileBox.FileName
BugFix = (Replace(lblFileName, "\\", "\"))
lblFileName.Caption = BugFix
End Sub

Private Sub FileBox_DblClick()
Shell ("Start ") & lblFileName.Caption
End Sub

Private Sub Form_Load()
DriveBox.Drive = CurDir
FileBox.Path = CurDir
DirBox.Path = CurDir
End Sub
