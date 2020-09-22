VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create your own file archives"
   ClientHeight    =   4095
   ClientLeft      =   2310
   ClientTop       =   1125
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6840
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   960
      Left            =   1680
      TabIndex        =   10
      Top             =   3045
      Width           =   4635
      Begin VB.PictureBox picProgress 
         Height          =   330
         Left            =   210
         ScaleHeight     =   270
         ScaleWidth      =   4155
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdExtractArchiv 
      Caption         =   "extract archive"
      Height          =   330
      Left            =   4095
      TabIndex        =   9
      Top             =   2415
      Width           =   2220
   End
   Begin VB.CommandButton cmdMakeArchiv 
      Caption         =   "create archive"
      Height          =   330
      Left            =   1680
      TabIndex        =   8
      Top             =   2415
      Width           =   2220
   End
   Begin VB.ComboBox cmbPattern 
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2415
      Width           =   1170
   End
   Begin VB.CommandButton cmdOrdner 
      Height          =   330
      Left            =   6405
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1650
      Width           =   330
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   6105
   End
   Begin VB.Label Label2 
      Caption         =   "File-Mask:"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   2145
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Please select a directory for extracting and saving of files:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1365
      Width           =   6420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "by IRNCHEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   225
      Index           =   2
      Left            =   105
      TabIndex        =   2
      Top             =   945
      UseMnemonic     =   0   'False
      Width           =   6630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "EASY AND CLEAN CODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   630
      UseMnemonic     =   0   'False
      Width           =   6630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CREATE YOUR OWN FILE ARCHIVES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      UseMnemonic     =   0   'False
      Width           =   6525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code shows, howto make your own archives!!!
'Clean code and full commented.
'P.S.: PLEASE VOTE FOR THIS CODE!!!!!

Private Sub cmdExtractArchiv_Click()
  Dim nFiles As Long
  Dim sPath As String
  
  ' select the archive
  On Local Error Resume Next
  With CommonDialog1
    .CancelError = True
    .Filter = "Archiv-Datei (*.dat)|*.dat"
    'Show the dialog
    .ShowOpen
    If Err = 0 Then
      ' select the directory
      sPath = BrowseForFolder("Please select the directory to extract to:")
      If sPath <> "" Then
        ' Show the progressbar
        picProgress.Visible = True
        'extract the files
        nFiles = ExtractFilesFromArchiv(.FileName, _
          sPath)
          'Hide the progressbar
        picProgress.Visible = False
        
        MsgBox nFiles & " files extracted to " & sPath, 64
      End If
    End If
  End With
End Sub

Private Sub cmdMakeArchiv_Click()
  ' select archive
  Dim nFiles As Long
  
  On Local Error Resume Next
  With CommonDialog1
    .CancelError = True
    .Filter = "Archive (*.dat)|*.dat"
    .DefaultExt = ".dat"
    'Show the save dialog
    .ShowSave
    If Err = 0 Then
    'Show the progressbar
      picProgress.Visible = True
      'Save the files to the archive
      nFiles = SaveFilesToArchiv(txtDir.Text, _
        .FileName, cmbPattern.Text)
        'Hide the progressbar
      picProgress.Visible = False
      'Tell the user, how many files are stored in the archive
      MsgBox nFiles & " stored in " & _
        .FileName, 64
    End If
  End With
End Sub

Private Sub cmdOrdner_Click()
  ' select directory
  Dim sOrdner As String
  
  sOrdner = BrowseForFolder("Select a directory:")
  If sOrdner <> "" Then
    txtDir.Text = sOrdner
  End If
End Sub

Private Sub Form_Load()
  ' file-mask
  cmbPattern.AddItem "*.*"
  cmbPattern.ListIndex = 0
  MsgBox "PLEASE VOTE!!!!!" & vbCrLf & "I worked so long on this code!", vbInformation, "VOTE!!!!"
End Sub

' Save files to the archive
Public Function SaveFilesToArchiv( _
  ByVal sPath As String, _
  ByVal sArchiv As String, _
  Optional ByVal sPattern As String = "*.*") As Long

  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim FileData As String
  Dim File() As String
  Dim nFiles As Long
  Dim I As Long
  Dim lngUBound As Long

  ' Add backslash to the path
  If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"

  ' Get all files in the directory
  nFiles = 0

  DirName = Dir(sPath & sPattern, vbNormal)
  While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
      nFiles = nFiles + 1
      'Get files
      If nFiles > lngUBound Then lngUBound = 2 * nFiles
      ReDim Preserve File(lngUBound)
      File(nFiles) = DirName
    End If
    DirName = Dir
  Wend
  ReDim Preserve File(nFiles)

  ' If archiv exists already, delete it
  If Dir(sArchiv) <> "" Then Kill sArchiv

  ' Now save all files to the archive
  F = FreeFile
  Open sArchiv For Binary As #F

  ' Set number of files
  Put #F, , nFiles

  For I = 1 To nFiles
    ' Save filename
    nLenFileName = Len(File(I))
    Put #F, , nLenFileName
    Put #F, , File(I)

    ' Read filedata
    n = FreeFile
    Open sPath + File(I) For Binary As #n
    FileData = Space$(LOF(n))
    Get #n, , FileData
    Close #n

    ' Save filedata to the archive
    nLenFileData = Len(FileData)
    Put #F, , nLenFileData
    Put #F, , FileData
    
    ' Progress
    ShowProgress picProgress, I, 1, nFiles
    DoEvents
  Next I
  Close #F
  
  SaveFilesToArchiv = nFiles
End Function

' extract all files to the outputfolder
Public Function ExtractFilesFromArchiv( _
  ByVal sArchiv As String, _
  ByVal sDestDir As String) As Long
  
  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim FileData As String
  Dim File As String
  Dim nFiles As Long
  Dim I As Long
  
  ' check if Archiv exists
  If Dir(sArchiv) = "" Then
    MsgBox "The archive does not exist!", 16
    Exit Function
  End If
  
  ' add backslash to the path
  If Right$(sDestDir, 1) <> "\" Then _
    sDestDir = sDestDir + "\"
  
  ' Open the archive
  F = FreeFile
  Open sArchiv For Binary As #F
  
  ' Get number of Icons in the archive
  Get #F, , nFiles
  
  For I = 1 To nFiles
    ' get original filenames
    Get #F, , nLenFileName
    File = Space$(nLenFileName)
    Get #F, , File
    
    ' Read filedata
    Get #F, , nLenFileData
    FileData = Space$(nLenFileData)
    Get #F, , FileData
    
    ' Save file in "DestDir"
    n = FreeFile
    Open sDestDir + File For Output As #n
    Print #n, FileData;
    Close #n
    
    ' Progress
    ShowProgress picProgress, I, 1, nFiles
    DoEvents
  Next I
  Close #F
  
  ExtractFilesFromArchiv = nFiles
End Function
' Progressbar
Private Sub ShowProgress(picProgress As PictureBox, _
  ByVal Value As Long, _
  ByVal Min As Long, _
  ByVal Max As Long, _
  Optional ByVal bShowProzent As Boolean = True)
  
  Dim pWidth As Long
  Dim intProz As Integer
  Dim strProz As String
  
  ' colors
  Const progBackColor = &HC00000
  Const progForeColor = vbBlack
  Const progForeColorHighlight = vbWhite
  
  ' set Values
  If Value < Min Then Value = Min
  If Value > Max Then Value = Max
  
  ' Prozentwert ausrechnen
  If Max > 0 Then
    intProz = Int(Value / Max * 100 + 0.5)
  Else
    intProz = 100
  End If
    
  With picProgress
    ' check if AutoReadraw=True
    If .AutoRedraw = False Then .AutoRedraw = True
    
    ' clear the picturebox
    picProgress.Cls
    
    If Value > 0 Then
    
      ' calculate barwidth
      pWidth = .ScaleWidth / 100 * intProz
      
      ' Show bar
      picProgress.Line (0, 0)-(pWidth, .ScaleHeight), _
        progBackColor, BF
        
      ' show percent
      If bShowProzent Then
        strProz = CStr(intProz) & " %"
        .CurrentX = (.ScaleWidth - .TextWidth(strProz)) / 2
        .CurrentY = (.ScaleHeight - .TextHeight(strProz)) / 2
      
        ' Foregroundcolor
        If pWidth >= .CurrentX Then
          .ForeColor = progForeColorHighlight
        Else
          .ForeColor = progForeColor
        End If
      
        picProgress.Print strProz
      End If
    End If
  End With
End Sub

Private Sub txtDir_Change()
  cmdMakeArchiv.Enabled = (txtDir.Text <> "")
End Sub

