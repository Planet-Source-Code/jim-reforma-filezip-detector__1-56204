VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileZIP Detector by Jim Reforma [virushacker23@yahoo.com]"
   ClientHeight    =   6720
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      Height          =   615
      Left            =   5040
      Picture         =   "frmMain.frx":476A
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   18
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdComentario 
      Caption         =   "&About"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Frame frmSelection 
      Caption         =   "Selected Information"
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   4095
      Begin VB.TextBox txtComp 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Compressed size"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Uncompressed size"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtCrc 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Cyclic Redundancy Check"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtTipo 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "File or folder"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblComp 
         Caption         =   "Compressed size:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblDesc 
         Caption         =   "Uncompressed size:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblCrc32 
         Caption         =   "CRC:"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblTipo 
         Caption         =   "Type:"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.FileListBox Files 
      BackColor       =   &H00D8E9EC&
      Height          =   2820
      Left            =   3240
      Pattern         =   "*.zip"
      TabIndex        =   7
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drives 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.DirListBox Dir 
      BackColor       =   &H00D8E9EC&
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.ListBox lstFiles 
      BackColor       =   &H00D8E9EC&
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label lblSize 
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblTam 
      Caption         =   "Total size:"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblFiles 
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblCant 
      Caption         =   "Files in archive:"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================
'Comments!! Imagination is more important than knowledge,
'           for knowledge is limited while imagination embraces
'           the entire world.""
'==================================================================
' Copyright Version (c), July 2000
' Jim Reforma [virushacker23@yahoo.com]

Public Comment As String, ListErrores As String
Dim File As String, LFHS As String, ECD As String, Crc() As String, OldDrive As String
Dim NameLong As Long, LongTotName As Long, NumArInFile As Long
Dim CompLong() As Double, DescLong() As Double
Dim HuboError As Boolean
Const Title = "FileZIP Detector by Jim Reforma [virushacker23@yahoo.com]"

Private Sub cmdComentario_Click()
 frmAbout.Show vbModal
End Sub

Private Sub Dir_Change()
  Files.Path = Dir.Path
End Sub

Private Sub Drives_Change()
  On Error GoTo ErrVer
  Dir.Path = Drives.Drive
  OldDrive = Drives.Drive
  Exit Sub
ErrVer:
  If Err.Number = 68 Then
    If MsgBox("This drive is not ready", vbCritical + vbRetryCancel) = vbRetry Then
      Drives_Change
    Else
      Drives.Drive = OldDrive
    End If
  End If
End Sub

Private Sub Files_Click()
  HuboError = False
  Accion = True
  ListaErrores = ""
  Me.Caption = Title + " - " + Files.FileName + " - Reading archive..."
  Erase Crc, CompLong, DescLong
  Comment = ""
  lstFiles.Clear
  Open Verificar(Dir.Path) + Files.FileName For Binary Access Read As #1
  LongTotName = 0
  NuMar = 1
  If Input$(4, 1) <> LFHS Then
    MsgBox "Local file signature of file no. " + Format$(NuMar) + " is missing." + vbCrLf + "Unable to continue", vbCritical
    HuboError = True
    GoTo Terminate
  End If
Rep:
  Seek #1, Loc(1) + 11
  ReDim Preserve Crc(1 To NuMar)
  ReDim Preserve CompLong(1 To NuMar)
  ReDim Preserve DescLong(1 To NuMar)
  Crc(NuMar) = Invert(LCase(Check(Hex$(Asc(Input$(1, 1)))) + Check(Hex$(Asc(Input$(1, 1)))) + Check(Hex$(Asc(Input$(1, 1)))) + Check(Hex$(Asc(Input$(1, 1))))))
  CompLong(NuMar) = DeConvert(Input$(4, 1))
  DescLong(NuMar) = DeConvert(Input$(4, 1))
  NameLong = DeConvert(Input$(2, 1) + Chr$(0) + Chr$(0))
  If NameLong > 255 Then AgregarError 3
  Seek #1, Loc(1) + 3
  LongTotName = LongTotName + NameLong
  lstFiles.AddItem Left(Input$(NameLong, 1), 255)
  Seek #1, Loc(1) + CompLong(NuMar) + 1
  Select Case Input$(4, 1)
    Case LFHS
      NuMar = NuMar + 1
      GoTo Rep
    Case ECD
    Case Else
      If MsgBox("Error in the ZIP file structure." + vbCrLf + "Do you want " + Title + " to continue scanning?", vbCritical + vbYesNo) = vbNo Then
        HuboError = True
        GoTo Terminate
      End If
  End Select
  Seek #1, Loc(1) + (NuMar * 46) + LongTotName + 5
  NumArInFile = DeConvert(Input$(2, 1) + Chr$(0) + Chr$(0))
  If NumArInFile <> NuMar Then AgregarError 1
  If NumArInFile <> DeConvert(Input$(2, 1) + Chr$(0) + Chr$(0)) Then AgregarError 2
  Seek #1, Loc(1) + 9
  Comment = Input$(DeConvert(Input$(2, 1) + Chr$(0) + Chr$(0)), 1)
  lblFiles.Caption = Format$(NuMar)
  lblSize.Caption = Format$(Int(LOF(1) / 1024)) + " KB"
  Close
  Me.Caption = Title + " - " + Files.FileName
  If Comment <> "" Then frmComentario.Show vbModal
Terminate:
  If HuboError = True Then Me.Caption = Title: lstFiles.Clear: lblSize.Caption = "": lblFiles.Caption = ""
  If ListaErrores <> "" Then If MsgBox(Title + " had problems when scanning." + vbCrLf + "Do you want to view the messages?", vbQuestion + vbYesNo) = vbYes Then Accion = False: frmComentario.Show vbModal
  txtDesc.Text = ""
  txtComp.Text = ""
  txtCrc.Text = ""
  Close
End Sub

Private Sub Form_Initialize()
  If App.PrevInstance = True Then
    If MsgBox("Another instance of " + Title + " is running." + vbCrLf + "Do you want to start another one?", vbQuestion + vbYesNo) = vbNo Then End
  End If
End Sub

Private Sub Form_Load()
  ECD = "PK" + Chr$(1) + Chr$(2)
  LFHS = "PK" + Chr$(3) + Chr$(4)
  Me.Caption = Title
  OldDrive = Drives.Drive
End Sub

Private Sub imgLogo_Click()
  frmAbout.Show vbModal
End Sub

Private Sub lstFiles_Click()
  If CompLong(lstFiles.ListIndex + 1) < 1024 Then
    txtComp.Text = Format$(CompLong(lstFiles.ListIndex + 1)) + " bytes"
  Else
    txtComp.Text = Format$(Int(CompLong(lstFiles.ListIndex + 1) / 1024)) + " KB"
  End If
  If DescLong(lstFiles.ListIndex + 1) < 1024 Then
    txtDesc.Text = Format$(DescLong(lstFiles.ListIndex + 1)) + " bytes"
  Else
    txtDesc.Text = Format$(Int(DescLong(lstFiles.ListIndex + 1) / 1024)) + " bytes"
  End If
  txtCrc.Text = Crc(lstFiles.ListIndex + 1)
  If Right$(lstFiles.List(lstFiles.ListIndex), 1) = "\" Or Right$(lstFiles.List(lstFiles.ListIndex), 1) = "/" Then
    txtTipo.Text = "Folder"
  Else
    txtTipo.Text = "File"
  End If
  Me.Caption = Title + " - " + Files.FileName + " (" + lstFiles.List(lstFiles.ListIndex) + ")"
End Sub

Function DeConvert(Cadena As String) As Double
  If Len(Cadena) <> 4 Then AgregarError 0: Exit Function
  If Asc(Mid$(Cadena, 4)) = 0 Then
    If Asc(Mid$(Cadena, 3)) = 0 Then
      If Asc(Mid$(Cadena, 2)) = 0 Then
        DeConvert = CDbl(Asc(Mid$(Cadena, 1)))
      Else
        DeConvert = CDbl(Asc(Mid$(Cadena, 2))) * 256 + CDbl(Asc(Mid$(Cadena, 1)))
      End If
    Else
      DeConvert = CDbl(Asc(Mid$(Cadena, 3))) * 256 * 256 + CDbl(Asc(Mid$(Cadena, 2))) * 256 + CDbl(Asc(Mid$(Cadena, 1)))
    End If
  Else
    DeConvert = CDbl(Asc(Mid$(Cadena, 4))) * 256 * 256 * 256 + CDbl(Asc(Mid$(Cadena, 3))) * 256 * 256 + CDbl(Asc(Mid$(Cadena, 2))) * 256 + CDbl(Asc(Mid$(Cadena, 1)))
  End If
End Function

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub


Private Sub mnuAyuda_Click()

End Sub

Function Check(Cadena As String) As String
  If Len(Cadena) = 1 Then
    Check = "0" + Cadena
  Else
    Check = Cadena
  End If
End Function

Function Verificar(Cadena As String) As String
  If Right$(Cadena, 1) = "\" Then
    Verificar = Cadena
  Else
    Verificar = Cadena + "\"
  End If
End Function

Function Invert(Cadena As String) As String
  For i = Len(Cadena) - 1 To 0 Step -2
    Invert = Invert + Mid$(Cadena, i, 2)
  Next i
End Function

Private Sub mnuAcerca_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuSalir_Click()
  End
End Sub

'==================================================================
'Comments!! Imagination is more important than knowledge,
'           for knowledge is limited while imagination embraces
'           the entire world.""
'==================================================================

