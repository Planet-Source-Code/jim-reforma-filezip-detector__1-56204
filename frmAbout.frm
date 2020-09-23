VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About FileZip 32bits Detector"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00D8E9EC&
      Height          =   615
      Left            =   120
      Picture         =   "frmAbout.frx":476A
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00D8E9EC&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame fraCopyright 
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         Caption         =   "Copyright ©2000 by JimSoft, Inc.    All rights reserved.  virushacker23@yahoo.com"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label lblIntro 
      Caption         =   "FileZip 32bits Detector"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================
'Comments!! Imagination is more important than knowledge,
'           for knowledge is limited while imagination embraces
'           the entire world.""
'==================================================================Public Comment As String, ListErrores As String
' Copyright Version (c), July 2000
' Jim Reforma [virushacker23@yahoo.com]

Private Sub cmdAceptar_Click()
  Unload Me
End Sub

'==================================================================
'Comments!! Imagination is more important than knowledge,
'           for knowledge is limited while imagination embraces
'           the entire world.""
'==================================================================

