VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "CodeDescriptor"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Description"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "File1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.FileListBox File1 
         Height          =   4185
         Left            =   120
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin RichTextLib.RichTextBox Text1 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7435
         _Version        =   393217
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":0038
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OldPath As String
Private Sub Dir1_Change()
Text1.SaveFile OldPath, rtfText
Text1.Text = ""

OldPath = Dir1.Path & "\" & "Readme.txt"
If Len(Dir(Dir1.Path & "\" & "Readme.txt")) Then
Text1.LoadFile Dir1.Path & "\" & "Readme.txt"
End If
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
Dim x
Dim strSearch As String
Dim tmpString As String
If KeyAscii = 115 Then
    strSearch = InputBox("Write in the directory you are looking for", "Search...")
        For Each x In ListSubDirs(Dir1.Path & "\")
        
            If InStr(1, UCase(x), UCase(strSearch)) Then
                Dir1.Path = Dir1.Path & "\" & x
                Exit Sub
            End If
        Next x
MsgBox "No match for " & strSearch, vbInformation, "Find..."
End If
        
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\VB" 'Root folder of your vb directory
End Sub

Function ListSubDirs(ByVal Path As String) As Variant
'Made by Alex Wolfe

    'returns an array of directory names
    On Error Resume Next
    Dim Count, Dirs(), i, DirName ' Declare variables.
    DirName = Dir(Path, vbDirectory) ' Get first directory name.
    Count = 0


    Do While Not DirName = ""
        ' A file or directory name was returned


        If Not DirName = "." And Not DirName = ".." Then
            ' Not a parent or current directory entr
            '     y so process it


            If GetAttr(Path & DirName) And vbDirectory Then
                ' This is a directory
                ' Increase the size of the array by one
                '     element
                ReDim Preserve Dirs(Count + 1)
                Dirs(Count) = DirName ' Add directory name to array
                Count = Count + 1 ' Increment counter.
            End If
        End If
        DirName = Dir ' Get another directory name.
    Loop
    ReDim Preserve Dirs(Count - 1) 'remove the last empty element
    ListSubDirs = Dirs()
End Function


